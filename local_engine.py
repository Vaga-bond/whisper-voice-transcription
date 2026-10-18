"""Transcription locale par Nemotron 3.5 ASR Streaming, le modèle de Handy.

Le moteur est transcribe.cpp (paquet `transcribe-cpp`), celui qu'utilise Handy : il tourne sur
la carte graphique via Vulkan quand elle est disponible, sinon sur le processeur. Le fichier du
modèle est celui que Handy a déjà téléchargé dans le cache Hugging Face.

Le son est transcrit pendant l'enregistrement : à l'arrêt, il ne reste que la fin de la phrase à
décoder. Toutes les opérations sur le modèle passent par un seul fil d'exécution, qui le charge à
la demande et le libère après IDLE_UNLOAD secondes sans usage, comme Handy.
"""
import glob
import importlib.util
import os
import queue
import threading
import time
from pathlib import Path

import numpy as np

MODEL_ID = "nemotron-3.5-local"
HF_REPO = "handy-computer/nemotron-3.5-asr-streaming-0.6b-gguf"
MODEL_FILE = "nemotron-3.5-asr-streaming-0.6b-Q8_0.gguf"
MODEL_URL = f"https://huggingface.co/{HF_REPO}/resolve/main/{MODEL_FILE}"
LOCAL_MODELS_DIR = Path(__file__).parent / "models"

MODEL_RATE = 16000
IDLE_UNLOAD = 600


def find_model():
    """Chemin du fichier .gguf : NEMOTRON_MODEL, sinon le cache de Handy, sinon models/."""
    forced = os.environ.get("NEMOTRON_MODEL")
    if forced:
        return forced if os.path.isfile(forced) else None
    hub = os.environ.get("HF_HUB_CACHE") or os.path.join(
        os.environ.get("HF_HOME") or os.path.expanduser("~/.cache/huggingface"), "hub")
    pattern = os.path.join(hub, "models--" + HF_REPO.replace("/", "--"), "snapshots", "*", MODEL_FILE)
    found = sorted(glob.glob(pattern)) + sorted(glob.glob(str(LOCAL_MODELS_DIR / "*.gguf")))
    return found[0] if found else None


def missing_requirement():
    """Message d'erreur si le moteur ou le modèle manque, sinon None."""
    absent = [p for p, mod in (("transcribe-cpp", "transcribe_cpp"), ("soxr", "soxr"))
              if importlib.util.find_spec(mod) is None]
    if absent:
        return ("Paquets manquants : %s.\nRelancez setup.bat ou installez-les avec "
                "« pip install -r requirements.txt »." % ", ".join(absent))
    if find_model() is None:
        return ("Modèle Nemotron introuvable.\n\nIl est présent si Handy l'a téléchargé. Sinon, "
                "placez le fichier %s dans le dossier « models » de l'application :\n%s"
                % (MODEL_FILE, MODEL_URL))
    return None


class NemotronEngine:
    """Une dictée à la fois : begin, feed au fil du son, puis finish (texte) ou cancel."""

    def __init__(self, language="fr-FR"):
        self.language = language
        self.device = None
        self._commands = queue.Queue()
        self._thread = None
        self._thread_lock = threading.Lock()

    def preload(self):
        self._send("preload", None)

    def begin(self, input_rate):
        self._send("begin", input_rate)

    def feed(self, data):
        """Son brut du micro : int16 mono, au taux annoncé dans begin."""
        self._commands.put(("audio", data))

    def cancel(self):
        self._commands.put(("cancel", None))

    def finish(self):
        """Termine la dictée en cours et renvoie son texte ; lève l'erreur du moteur s'il y en a eu une."""
        reply = queue.Queue(maxsize=1)
        self._send("end", reply)
        ok, value = reply.get()
        if not ok:
            raise value
        return value

    def _send(self, kind, arg):
        with self._thread_lock:
            if self._thread is None or not self._thread.is_alive():
                self._thread = threading.Thread(target=self._run, daemon=True)
                self._thread.start()
        self._commands.put((kind, arg))

    def _run(self):
        model = session = stream = resampler = None
        error = None
        last_used = time.monotonic()

        def load():
            nonlocal model, session
            if model is not None:
                return
            path = find_model()
            if path is None:
                raise FileNotFoundError(missing_requirement())
            import transcribe_cpp
            model = transcribe_cpp.Model(path)
            session = model.session().__enter__()
            self.device = model.device.description
            # Premier passage à vide : compile les noyaux GPU avant la première vraie dictée.
            with session.stream(language=self.language) as warmup:
                warmup.feed(np.zeros(MODEL_RATE, dtype=np.float32))
                warmup.finalize()

        def unload():
            nonlocal model, session
            if session is not None:
                session.__exit__(None, None, None)
            if model is not None:
                model.close()
            model = session = None

        def close_stream():
            nonlocal stream, resampler
            if stream is not None:
                stream.__exit__(None, None, None)
            stream = resampler = None

        pending = None
        while True:
            if pending is not None:
                (kind, arg), pending = pending, None
            else:
                try:
                    kind, arg = self._commands.get(timeout=30)
                except queue.Empty:
                    if model is not None and stream is None and time.monotonic() - last_used > IDLE_UNLOAD:
                        unload()
                    continue
            last_used = time.monotonic()

            # Le micro livre de petits blocs : on donne d'un coup au modèle tout le son en attente.
            if kind == "audio":
                chunks = [arg]
                while pending is None:
                    try:
                        nxt = self._commands.get_nowait()
                    except queue.Empty:
                        break
                    if nxt[0] == "audio":
                        chunks.append(nxt[1])
                    else:
                        pending = nxt
                arg = b"".join(chunks)

            try:
                if kind == "preload":
                    load()
                elif kind == "begin":
                    close_stream()
                    error = None
                    load()
                    import soxr
                    resampler = soxr.ResampleStream(arg, MODEL_RATE, 1, dtype="float32")
                    stream = session.stream(language=self.language).__enter__()
                elif kind == "audio":
                    if stream is not None:
                        pcm = np.frombuffer(arg, dtype="<i2").astype(np.float32) / 32768.0
                        out = resampler.resample_chunk(pcm)
                        if len(out):
                            stream.feed(out)
                elif kind == "cancel":
                    close_stream()
                elif kind == "end":
                    if error is not None:
                        arg.put((False, error))
                    elif stream is None:
                        arg.put((True, ""))
                    else:
                        tail = resampler.resample_chunk(np.zeros(0, dtype=np.float32), last=True)
                        if len(tail):
                            stream.feed(tail)
                        stream.finalize()
                        text = stream.text().full.strip()
                        close_stream()
                        arg.put((True, text))
                    error = None
            except Exception as e:
                close_stream()
                if kind == "end":
                    arg.put((False, e))
                elif kind != "preload":
                    error = e
                print(f"⚠️ Nemotron ({kind}) : {e}")
