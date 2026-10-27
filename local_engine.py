"""Transcription locale par transcribe.cpp (paquet `transcribe-cpp`), le moteur de Handy.

Il tourne sur la carte graphique via Vulkan quand elle est disponible, sinon sur le processeur.
Chaque modèle est un fichier .gguf, lu dans le cache Hugging Face quand Handy l'a téléchargé,
sinon dans le dossier « models » de l'application.

Nemotron transcrit pendant l'enregistrement : à l'arrêt, il ne reste que la fin de la phrase à
décoder. Les autres modèles ne savent pas transcrire en continu : ils reçoivent le son au fil de
l'eau et le transcrivent d'un bloc à l'arrêt. Toutes les opérations passent par un seul fil
d'exécution, qui charge le modèle demandé (en libérant le précédent) et le libère après
IDLE_UNLOAD secondes sans usage, comme Handy.
"""
import glob
import importlib.util
import os
import queue
import threading
import time
import urllib.request
from pathlib import Path
from typing import NamedTuple

import numpy as np


class LocalModel(NamedTuple):
    label: str
    repo: str
    file: str
    size_mb: int
    streaming: bool = False

    @property
    def url(self):
        return f"https://huggingface.co/{self.repo}/resolve/main/{self.file}"

    @property
    def size_text(self):
        if self.size_mb < 1000:
            return f"{self.size_mb} Mo"
        return f"{self.size_mb / 1000:.1f} Go".replace(".", ",")


# Identifiant (préférences, historique) → modèle. Q8_0 partout : la quantification la plus
# basse qui reste au niveau du modèle d'origine.
LOCAL_MODELS = {
    "whisper-large-v3-turbo-local": LocalModel(
        "Whisper large-v3 turbo local (recommandé)", "handy-computer/whisper-large-v3-turbo-gguf",
        "whisper-large-v3-turbo-Q8_0.gguf", 886),
    "nemotron-3.5-local": LocalModel(
        "Nemotron 3.5 local (Handy)", "handy-computer/nemotron-3.5-asr-streaming-0.6b-gguf",
        "nemotron-3.5-asr-streaming-0.6b-Q8_0.gguf", 720, streaming=True),
    "qwen3-asr-1.7b-local": LocalModel(
        "Qwen3-ASR 1.7B local", "handy-computer/Qwen3-ASR-1.7B-gguf",
        "Qwen3-ASR-1.7B-Q8_0.gguf", 2185),
    "granite-speech-4.1-2b-local": LocalModel(
        "Granite Speech 4.1 2B local", "handy-computer/granite-speech-4.1-2b-gguf",
        "granite-speech-4.1-2b-Q8_0.gguf", 2560),
    "parakeet-tdt-0.6b-v3-local": LocalModel(
        "Parakeet TDT 0.6B v3 local", "handy-computer/parakeet-tdt-0.6b-v3-gguf",
        "parakeet-tdt-0.6b-v3-Q8_0.gguf", 740),
}
LOCAL_MODELS_DIR = Path(__file__).parent / "models"

MODEL_RATE = 16000
IDLE_UNLOAD = 600


def find_model(model_id):
    """Chemin du fichier .gguf du modèle : cache Hugging Face (Handy), sinon models/."""
    spec = LOCAL_MODELS[model_id]
    hub = os.environ.get("HF_HUB_CACHE") or os.path.join(
        os.environ.get("HF_HOME") or os.path.expanduser("~/.cache/huggingface"), "hub")
    pattern = os.path.join(hub, "models--" + spec.repo.replace("/", "--"), "snapshots", "*", spec.file)
    found = sorted(glob.glob(pattern))
    local = LOCAL_MODELS_DIR / spec.file
    if local.is_file():
        found.append(str(local))
    return found[0] if found else None


def missing_packages():
    """Message d'erreur si le moteur n'est pas installé, sinon None."""
    absent = [p for p, mod in (("transcribe-cpp", "transcribe_cpp"), ("soxr", "soxr"))
              if importlib.util.find_spec(mod) is None]
    if absent:
        return ("Paquets manquants : %s.\nRelancez setup.bat ou installez-les avec "
                "« pip install -r requirements.txt »." % ", ".join(absent))
    return None


def missing_requirement(model_id):
    """Message d'erreur si le moteur ou le fichier du modèle manque, sinon None."""
    problem = missing_packages()
    if problem:
        return problem
    if find_model(model_id) is None:
        spec = LOCAL_MODELS[model_id]
        return ("Modèle %s introuvable.\n\nPlacez le fichier %s dans le dossier « models » de "
                "l'application :\n%s" % (spec.label, spec.file, spec.url))
    return None


def download(model_id, progress):
    """Télécharge le fichier du modèle dans models/ ; progress(reçus, total) au fil de l'eau."""
    spec = LOCAL_MODELS[model_id]
    LOCAL_MODELS_DIR.mkdir(exist_ok=True)
    target = LOCAL_MODELS_DIR / spec.file
    part = target.with_name(target.name + ".part")
    try:
        with urllib.request.urlopen(spec.url, timeout=30) as response, open(part, "wb") as out:
            total = int(response.headers.get("Content-Length") or 0)
            done = 0
            while True:
                block = response.read(1 << 20)
                if not block:
                    break
                out.write(block)
                done += len(block)
                progress(done, total)
        if total and done != total:
            raise IOError(f"téléchargement incomplet ({done} octets sur {total})")
        part.replace(target)
    except BaseException:
        part.unlink(missing_ok=True)
        raise


def _language_code(model, language):
    """Code de langue tel que le modèle l'attend (« fr », « fr-FR »…), None s'il détecte seul."""
    codes = model.capabilities.languages
    for code in codes:
        if code == language:
            return code
    for code in codes:
        if code.split("-")[0] == language:
            return code
    return None


class LocalEngine:
    """Une dictée à la fois : begin, feed au fil du son, puis finish (texte) ou cancel."""

    def __init__(self, language="fr"):
        self.language = language
        self.device = None
        self._commands = queue.Queue()
        self._thread = None
        self._thread_lock = threading.Lock()

    def preload(self, model_id):
        self._send("preload", model_id)

    def begin(self, model_id, input_rate):
        self._send("begin", (model_id, input_rate))

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
        model = session = None
        loaded_id = language = None
        # Dictée en cours : `stream` pour un modèle en continu, `chunks` (son déjà au taux du
        # modèle) pour les autres.
        resampler = stream = chunks = None
        error = None
        last_used = time.monotonic()

        def load(model_id):
            nonlocal model, session, loaded_id, language
            if loaded_id == model_id:
                return
            unload()
            path = find_model(model_id)
            if path is None:
                raise FileNotFoundError(missing_requirement(model_id))
            import transcribe_cpp
            model = transcribe_cpp.Model(path)
            session = model.session().__enter__()
            loaded_id = model_id
            language = _language_code(model, self.language)
            self.device = model.device.description
            # Premier passage à vide : compile les noyaux GPU avant la première vraie dictée.
            silence = np.zeros(MODEL_RATE, dtype=np.float32)
            if LOCAL_MODELS[model_id].streaming:
                with session.stream(language=language) as warmup:
                    warmup.feed(silence)
                    warmup.finalize()
            else:
                session.run(silence, language=language, timestamps="none")

        def unload():
            nonlocal model, session, loaded_id
            if session is not None:
                session.__exit__(None, None, None)
            if model is not None:
                model.close()
            model = session = loaded_id = None

        def close_dictation():
            nonlocal stream, chunks, resampler
            if stream is not None:
                stream.__exit__(None, None, None)
            stream = chunks = resampler = None

        def transcribe_rest():
            """Fin de la dictée : décode ce qui reste et renvoie le texte complet."""
            tail = resampler.resample_chunk(np.zeros(0, dtype=np.float32), last=True)
            if stream is not None:
                if len(tail):
                    stream.feed(tail)
                stream.finalize()
                return stream.text().full.strip()
            if len(tail):
                chunks.append(tail)
            if not chunks:
                return ""
            pcm = np.concatenate(chunks)
            return session.run(pcm, language=language, timestamps="none").text.strip()

        pending = None
        while True:
            if pending is not None:
                (kind, arg), pending = pending, None
            else:
                try:
                    kind, arg = self._commands.get(timeout=30)
                except queue.Empty:
                    idle = stream is None and chunks is None
                    if model is not None and idle and time.monotonic() - last_used > IDLE_UNLOAD:
                        unload()
                    continue
            last_used = time.monotonic()

            # Le micro livre de petits blocs : on traite d'un coup tout le son en attente.
            if kind == "audio":
                parts = [arg]
                while pending is None:
                    try:
                        nxt = self._commands.get_nowait()
                    except queue.Empty:
                        break
                    if nxt[0] == "audio":
                        parts.append(nxt[1])
                    else:
                        pending = nxt
                arg = b"".join(parts)

            try:
                if kind == "preload":
                    load(arg)
                elif kind == "begin":
                    model_id, input_rate = arg
                    close_dictation()
                    error = None
                    load(model_id)
                    import soxr
                    resampler = soxr.ResampleStream(input_rate, MODEL_RATE, 1, dtype="float32")
                    if LOCAL_MODELS[model_id].streaming:
                        stream = session.stream(language=language).__enter__()
                    else:
                        chunks = []
                elif kind == "audio":
                    if resampler is not None:
                        pcm = np.frombuffer(arg, dtype="<i2").astype(np.float32) / 32768.0
                        out = resampler.resample_chunk(pcm)
                        if len(out):
                            if stream is not None:
                                stream.feed(out)
                            else:
                                chunks.append(out)
                elif kind == "cancel":
                    close_dictation()
                elif kind == "end":
                    if error is not None:
                        arg.put((False, error))
                    elif resampler is None:
                        arg.put((True, ""))
                    else:
                        text = transcribe_rest()
                        close_dictation()
                        arg.put((True, text))
                    error = None
            except Exception as e:
                close_dictation()
                if kind == "end":
                    arg.put((False, e))
                elif kind != "preload":
                    error = e
                print(f"⚠️ Modèle local ({kind}) : {e}")
