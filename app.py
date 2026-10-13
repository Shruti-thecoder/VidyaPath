from __future__ import annotations

import re
import os
import json
import base64
import hashlib
import textwrap
import array
import math
import sys
import tempfile
import wave
from html import escape
from concurrent.futures import ThreadPoolExecutor
from dataclasses import dataclass
from datetime import datetime, timezone
from io import BytesIO
from urllib.parse import parse_qs, urlencode, urlparse
from urllib.error import HTTPError, URLError
from urllib.request import Request, urlopen
from uuid import uuid4

import streamlit as st
from google import genai
from google.genai import types
from pypdf import PdfReader
from pptx import Presentation
from supabase import create_client


@dataclass
class Passage:
    filename: str
    location: str
    text: str


def get_secret(name: str, default: str = "") -> str:
    value = os.getenv(name, "").strip()
    if value:
        return value
    try:
        return str(st.secrets.get(name, default)).strip()
    except Exception:
        return default


def get_supabase_client():
    if "supabase_client" not in st.session_state:
        st.session_state.supabase_client = create_client(
            get_secret("SUPABASE_URL"), get_secret("SUPABASE_ANON_KEY")
        )
    return st.session_state.supabase_client


def load_quiz_history(client) -> list[dict]:
    response = client.table("quiz_results").select("topic,score,total,difficulty,created_at").order("created_at").execute()
    return response.data or []


def load_video_sources(client, user_id: str) -> list[dict]:
    response = client.table("video_sources").select("video_id,title,url,segments,created_at").eq("user_id", user_id).order("created_at").execute()
    return response.data or []


def save_video_source(client, user_id: str, source: dict) -> None:
    client.table("video_sources").upsert(
        {"user_id": user_id, **source}, on_conflict="user_id,video_id"
    ).execute()


def delete_video_source(client, user_id: str, video_id: str) -> None:
    client.table("video_sources").delete().eq("user_id", user_id).eq("video_id", video_id).execute()


def save_quiz_result(client, user_id: str, result: dict) -> None:
    client.table("quiz_results").insert({"user_id": user_id, **result}).execute()


def load_tutor_history(client, user_id: str, limit: int = 50) -> list[dict]:
    response = (
        client.table("tutor_history")
        .select("id,question,answer,sources,created_at")
        .eq("user_id", user_id)
        .order("created_at", desc=True)
        .limit(limit)
        .execute()
    )
    return response.data or []


def save_tutor_history(client, user_id: str, entry: dict) -> None:
    client.table("tutor_history").insert({"user_id": user_id, **entry}).execute()


def load_study_notes(client, user_id: str, limit: int = 100) -> list[dict]:
    response = (
        client.table("study_notes")
        .select("id,source_filename,source_location,note,created_at")
        .eq("user_id", user_id)
        .order("created_at", desc=True)
        .limit(limit)
        .execute()
    )
    return response.data or []


def save_study_note(client, user_id: str, entry: dict) -> dict:
    response = (
        client.table("study_notes")
        .insert({
            "user_id": user_id,
            "source_filename": entry["source_filename"],
            "source_location": entry["source_location"],
            "note": entry["note"],
        })
        .select("id,source_filename,source_location,note,created_at")
        .execute()
    )
    if not response.data:
        raise RuntimeError("The note wasn't returned after saving. Please try again.")
    return response.data[0]


def delete_study_note(client, user_id: str, note_id: int) -> None:
    client.table("study_notes").delete().eq("user_id", user_id).eq("id", note_id).execute()


def make_tutor_history_entry(question: str, answer: str, passages: list[Passage]) -> dict:
    return {
        "question": question,
        "answer": answer,
        "sources": [
            {"filename": passage.filename, "location": passage.location, "text": passage.text}
            for passage in passages
        ],
        "created_at": datetime.now(timezone.utc).isoformat(),
    }


def remember_tutor_interaction(entry: dict, client, user_id: str | None) -> None:
    st.session_state.tutor_history.insert(0, entry)
    if user_id:
        try:
            save_tutor_history(client, user_id, entry)
            st.session_state.tutor_history_error = None
        except Exception:
            st.session_state.tutor_history_error = (
                "Your answer is available for this session, but account history storage isn't set up yet. "
                "Run supabase_tutor_history.sql in the Supabase SQL Editor to keep it after you return."
            )


def load_used_questions(client, user_id: str, topic_key: str) -> list[str]:
    response = client.table("used_questions").select("question").eq("user_id", user_id).eq("topic_key", topic_key).execute()
    return [row["question"] for row in (response.data or [])]


def save_used_questions(client, user_id: str, topic_key: str, questions: list[str]) -> None:
    if questions:
        client.table("used_questions").insert([
            {"user_id": user_id, "topic_key": topic_key, "question": question}
            for question in questions
        ]).execute()


STOP_WORDS = {
    "a", "an", "and", "are", "as", "at", "be", "did", "do", "does", "for", "from",
    "give", "how", "i", "in", "is", "it", "me", "of", "on", "or", "please", "the",
    "to", "was", "what", "when", "where", "which", "who", "why", "with", "you",
}


def split_passage(text: str, filename: str, location: str, size: int = 1000) -> list[Passage]:
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    lines = [re.sub(r"[ \t]+", " ", line).strip() for line in text.splitlines()]
    lines = [line for line in lines if line]
    if not lines and text.strip():
        lines = [re.sub(r"\s+", " ", text).strip()]
    if not text:
        return []
    chunks: list[str] = []
    current = ""
    for line in lines:
        while len(line) > size:
            if current:
                chunks.append(current)
                current = ""
            chunks.append(line[:size])
            line = line[size:]
        if current and len(current) + len(line) + 1 > size:
            chunks.append(current)
            current = line
        else:
            current = f"{current}\n{line}" if current else line
    if current:
        chunks.append(current)
    return [Passage(filename, location, chunk) for chunk in chunks]


def read_file(filename: str, file_bytes: bytes) -> list[Passage]:
    data = BytesIO(file_bytes)
    passages: list[Passage] = []
    if filename.lower().endswith(".pdf"):
        reader = PdfReader(data)
        for page_number, page in enumerate(reader.pages, start=1):
            passages.extend(split_passage(page.extract_text() or "", filename, f"Page {page_number}"))
    elif filename.lower().endswith(".pptx"):
        deck = Presentation(data)
        for slide_number, slide in enumerate(deck.slides, start=1):
            parts = [shape.text for shape in slide.shapes if getattr(shape, "has_text_frame", False) and shape.text.strip()]
            passages.extend(split_passage("\n".join(parts), filename, f"Slide {slide_number}"))
    return passages


def transcribe_handwritten_pdf(file_bytes: bytes, api_key: str, filename: str) -> list[Passage]:
    """Transcribe a scanned PDF through Gemini only after the user opts in."""
    if len(file_bytes) > 50 * 1024 * 1024:
        raise ValueError("Gemini can transcribe PDFs up to 50 MB. This file is larger.")
    pdf_part = types.Part.from_bytes(data=file_bytes, mime_type="application/pdf")
    # Keep a strong reference to the client for the full synchronous request. The
    # short-lived chained form can leave the underlying HTTP client closed on reruns.
    gemini_client = genai.Client(api_key=api_key)
    try:
        response = gemini_client.models.generate_content(
            model="gemini-3.5-flash-lite",
            contents=[
                pdf_part,
                "Transcribe the handwritten text in this PDF as accurately as possible. "
                "Do not summarize or add explanations. Preserve headings, lists, and meaningful "
                "line breaks. Put each page under an exact marker like [[PAGE 1]], [[PAGE 2]]. "
                "Write [unclear] where handwriting cannot be read. Return only the transcription.",
            ],
        )
    finally:
        gemini_client.close()
    text = (response.text or "").strip()
    if not text:
        raise ValueError("Gemini couldn't read handwriting from this PDF. Try a clearer scan.")
    page_sections = re.split(r"\[\[PAGE\s+(\d+)\]\]", text, flags=re.IGNORECASE)
    passages: list[Passage] = []
    if len(page_sections) > 1:
        for index in range(1, len(page_sections), 2):
            page_number = page_sections[index]
            page_text = page_sections[index + 1] if index + 1 < len(page_sections) else ""
            passages.extend(split_passage(page_text, filename, f"Page {page_number}"))
    else:
        passages = split_passage(text, filename, "Transcribed PDF")
    if not passages:
        raise ValueError("Gemini couldn't find readable handwriting in this PDF.")
    return passages


def get_youtube_video_id(url: str) -> str | None:
    parsed = urlparse(url.strip())
    host = parsed.netloc.lower().split(":", 1)[0]
    if host in {"youtu.be", "www.youtu.be"}:
        video_id = parsed.path.strip("/").split("/")[0]
    elif host in {"youtube.com", "www.youtube.com", "m.youtube.com", "music.youtube.com"}:
        if parsed.path == "/watch":
            video_id = parse_qs(parsed.query).get("v", [""])[0]
        else:
            parts = [part for part in parsed.path.split("/") if part]
            video_id = parts[1] if len(parts) > 1 and parts[0] in {"embed", "shorts", "live"} else ""
    else:
        return None
    return video_id if re.fullmatch(r"[A-Za-z0-9_-]{11}", video_id) else None


def _friendly_network_error(error: Exception) -> str | None:
    """Turn raw DNS/connection/timeout failures into advice a learner can act on."""
    text = f"{type(error).__name__} {error}".lower()
    if any(key in text for key in ("getaddrinfo", "11001", "name resolution", "connecterror", "connection refused",
                                   "network is unreachable", "connection reset", "connecttimeout")):
        return ("VidyaPath can't reach Google Gemini from this computer. Check your internet connection, turn off any "
                "VPN or proxy, or try another network (changing your DNS to 8.8.8.8 often fixes this), then try again.")
    if "timeout" in text or "timed out" in text or "deadline" in text:
        return "Google Gemini took too long to answer. Try a shorter video (under about 30 minutes) or try again."
    return None


def parse_video_timestamp(value: str) -> int | None:
    """Read 'MM:SS', 'H:MM:SS' or 'HH:MM:SS' (also tolerates '[12:30]' or '12:30.5') into whole seconds."""
    match = re.fullmatch(r"\[?\s*(?:(\d{1,3}):)?(\d{1,3}):(\d{1,2})(?:[.,]\d+)?\s*\]?", value.strip())
    if not match:
        return None
    hours, minutes, seconds = int(match.group(1) or 0), int(match.group(2)), int(match.group(3))
    if seconds > 59 or (match.group(1) is not None and minutes > 59):
        return None
    return hours * 3600 + minutes * 60 + seconds


def build_video_outline(raw: str, video_id: str, url: str, snap_to: list[int] | None = None) -> dict:
    """Parse Gemini's JSON outline. With snap_to (real caption start times), every timestamp is moved to the nearest one."""
    try:
        payload = json.loads(raw)
    except json.JSONDecodeError:
        start, end = raw.find("{"), raw.rfind("}")
        if start < 0 or end <= start:
            raise ValueError("The video outline could not be read. Please try again.")
        payload = json.loads(raw[start : end + 1])
    if isinstance(payload, list) and payload and isinstance(payload[0], dict):
        payload = payload[0]
    if not isinstance(payload, dict):
        raise ValueError("The video outline came back in an unexpected format.")
    clean_segments = []
    last_seconds = -1
    for segment in payload.get("segments", []):
        if not isinstance(segment, dict):
            continue
        seconds = parse_video_timestamp(str(segment.get("timestamp", "")))
        topic = str(segment.get("topic", "")).strip()
        notes = str(segment.get("notes", "")).strip()
        if seconds is not None and snap_to:
            seconds = min(snap_to, key=lambda moment: abs(moment - seconds))
        # Accept MM:SS, H:MM:SS or HH:MM:SS, then keep timestamps in strictly increasing order.
        if seconds is not None and seconds > last_seconds and topic and notes:
            last_seconds = seconds
            clean_segments.append({
                "timestamp": f"{seconds // 3600:02d}:{seconds % 3600 // 60:02d}:{seconds % 60:02d}",
                "topic": topic[:120],
                "notes": notes[:700],
            })
    if not clean_segments:
        raise ValueError("Gemini couldn't find timestamped topics in this video. Check that it is public and try again.")
    title = str(payload.get("title", "")).strip()[:120] or f"YouTube lecture {video_id}"
    return {"video_id": video_id, "title": title, "url": url, "segments": clean_segments}


def fetch_youtube_transcript(video_id: str) -> list[tuple[float, str]] | None:
    """Fetch the video's existing captions (with exact start times) in a few seconds. None if unavailable."""
    try:
        from youtube_transcript_api import YouTubeTranscriptApi
    except ImportError:
        return None

    def fetch() -> list[tuple[float, str]]:
        api = YouTubeTranscriptApi()
        try:
            fetched = api.fetch(video_id, languages=["en", "en-IN", "en-US", "en-GB", "hi"])
        except Exception:
            # Preferred languages missing: use whichever caption track the video has.
            fetched = next(iter(api.list(video_id))).fetch()
        return [(float(item.start), str(item.text)) for item in fetched]

    executor = ThreadPoolExecutor(max_workers=1)
    try:
        rows = executor.submit(fetch).result(timeout=20)
    except Exception:
        return None
    finally:
        executor.shutdown(wait=False)
    rows = [(start, re.sub(r"\s+", " ", text).strip()) for start, text in rows if text.strip()]
    return rows or None


def group_transcript(rows: list[tuple[float, str]]) -> list[tuple[int, str]]:
    """Merge caption snippets into ~20-second lines so long videos stay small and quick to process."""
    span = rows[-1][0] - rows[0][0]
    window = max(20.0, span / 300)
    lines: list[tuple[int, str]] = []
    line_start, words = rows[0][0], []
    for start, text in rows:
        if words and start - line_start >= window:
            lines.append((int(line_start), " ".join(words)))
            line_start, words = start, []
        words.append(text)
    if words:
        lines.append((int(line_start), " ".join(words)))
    return lines


def analyze_youtube_video(video_id: str, api_key: str) -> dict:
    """Fast path: outline from the video's captions (seconds, exact times). Fallback: Gemini watches the video."""
    url = f"https://www.youtube.com/watch?v={video_id}"
    rows = fetch_youtube_transcript(video_id)
    if rows:
        lines = group_transcript(rows)
        transcript_text = "\n".join(f"[{s // 60}:{s % 60:02d}] {text}" for s, text in lines)
        prompt = (
            "Below is the timestamped transcript of a lecture video; each line starts with its time as [M:SS]. "
            "Create a concise, useful outline of its main topics and meaningful subtopics in chronological order, "
            "usually 5–20 segments. Return JSON only in this shape: "
            '{"title":"short accurate title","segments":[{"timestamp":"M:SS","topic":"short topic heading",'
            '"notes":"one or two concise factual sentences"}]}. '
            "Each timestamp must be copied exactly from the [M:SS] marker of the line where that topic begins, "
            "and timestamps must strictly increase. Use only the transcript; do not invent topics.\n\n" + transcript_text
        )
        client = genai.Client(api_key=api_key, http_options=types.HttpOptions(timeout=120_000))
        try:
            response = client.models.generate_content(
                model="gemini-3.5-flash-lite",
                contents=prompt,
                config=types.GenerateContentConfig(response_mime_type="application/json", temperature=0.2),
            )
            raw = re.sub(r"^```(?:json)?\s*|\s*```$", "", (response.text or "").strip(), flags=re.IGNORECASE)
            return build_video_outline(raw, video_id, url, snap_to=[s for s, _ in lines])
        except Exception as error:
            friendly = _friendly_network_error(error)
            if friendly:
                raise ConnectionError(friendly) from error
            # Anything else (unreadable answer, rejected option): fall through to watching the video instead.
        finally:
            client.close()
    return analyze_video_with_gemini(video_id, api_key)


def analyze_video_with_gemini(video_id: str, api_key: str) -> dict:
    url = f"https://www.youtube.com/watch?v={video_id}"
    prompt = (
        "Analyze this public lecture video and create a concise, useful timestamped outline. "
        "Return JSON only in this shape: "
        '{"title":"short accurate title","segments":[{"spoken_start":"the first 6-10 words the speaker says at that moment",'
        '"timestamp":"MM:SS","topic":"short topic heading","notes":"one or two concise factual sentences"}]}. '
        "Each timestamp must be the exact moment the speaker BEGINS that topic, measured from the very start of the video "
        "(0:00). Write minutes and seconds as MM:SS, and use H:MM:SS only after the video passes one hour. "
        "First find the words spoken at that moment and copy them into spoken_start, then give the timestamp where those words "
        "are heard; never estimate evenly spaced times. Timestamps must strictly increase. "
        "Include the main topics and meaningful subtopics in chronological order, usually 5–20 segments. "
        "Do not invent topics or times. Do not include a general summary outside the JSON."
    )
    # Speed: low frame resolution (lectures are carried by the audio) and JSON output. Frame rate is left at
    # the default because sampling fewer frames makes timestamps less precise. A timeout stops endless hanging.
    fast_part = types.Part(file_data=types.FileData(file_uri=url))
    fast_config = types.GenerateContentConfig(
        response_mime_type="application/json",
        media_resolution=types.MediaResolution.MEDIA_RESOLUTION_LOW,
        temperature=0.2,
    )
    client = genai.Client(api_key=api_key, http_options=types.HttpOptions(timeout=240_000))
    try:
        try:
            response = client.models.generate_content(
                model="gemini-3.5-flash-lite", contents=[fast_part, prompt], config=fast_config
            )
        except Exception as error:
            friendly = _friendly_network_error(error)
            if friendly:
                raise ConnectionError(friendly) from error
            if getattr(error, "code", None) != 400:
                raise
            # The model rejected a speed option: retry once with the plain request.
            response = client.models.generate_content(
                model="gemini-3.5-flash-lite",
                contents=[types.Part(file_data=types.FileData(file_uri=url)), prompt],
            )
        raw = re.sub(r"^```(?:json)?\s*|\s*```$", "", (response.text or "").strip(), flags=re.IGNORECASE)
    finally:
        client.close()
    return build_video_outline(raw, video_id, url)


def video_timestamp_url(source: dict, timestamp: str) -> str:
    hours, minutes, seconds = (int(part) for part in timestamp.split(":"))
    total_seconds = hours * 3600 + minutes * 60 + seconds
    return f"{source['url']}&{urlencode({'t': total_seconds})}"


def transcribe_voice_question(audio_bytes: bytes, mime_type: str, api_key: str) -> str:
    suffix = ".wav" if mime_type == "audio/wav" else ".audio"
    audio_path = None
    uploaded_audio = None
    client = genai.Client(api_key=api_key)
    try:
        with tempfile.NamedTemporaryFile(prefix="vidyapath_voice_", suffix=suffix, delete=False) as audio_file:
            audio_file.write(audio_bytes)
            audio_path = audio_file.name
        uploaded_audio = client.files.upload(file=audio_path)
        question = ""
        transcription_error = None
        try:
            response = client.models.generate_content(
                model="gemini-3.5-transcribe",
                contents=[uploaded_audio],
            )
            question = (response.text or "").strip().strip('"')
        except Exception as error:
            transcription_error = error

        # The dedicated transcriber can occasionally return an empty candidate
        # for microphone WAVs. Retry with the general audio-capable model and a
        # strict transcript-only instruction instead of failing the voice flow.
        if not question:
            try:
                fallback = client.models.generate_content(
                    model="gemini-3.5-flash-lite",
                    contents=[
                        uploaded_audio,
                        "Transcribe only the words actually spoken in this audio. "
                        "Preserve the speaker's language and wording. Do not answer, translate, "
                        "summarize, or guess. If there is no intelligible speech, return exactly [NO_SPEECH].",
                    ],
                )
                question = (fallback.text or "").strip().strip('"')
            except Exception as fallback_error:
                if transcription_error is not None:
                    raise RuntimeError(
                        f"Both transcription attempts failed ({transcription_error}; {fallback_error})."
                    ) from fallback_error
                raise
    finally:
        if uploaded_audio is not None:
            try:
                client.files.delete(name=uploaded_audio.name)
            except Exception:
                pass
        try:
            client.close()
        finally:
            if audio_path:
                try:
                    os.remove(audio_path)
                except OSError:
                    pass
    if not question or question.casefold() in {"[no speech]", "no speech", "[silence]", "silence"}:
        raise ValueError("The recording arrived, but the transcription service returned no words. Try a short 3–10 second question, speak clearly, then preview it before retrying.")
    return question


def recording_contains_speech(audio_bytes: bytes) -> bool:
    """Reject silent or near-silent WAV recordings before sending them for transcription."""
    try:
        with wave.open(BytesIO(audio_bytes), "rb") as recording:
            if recording.getcomptype() != "NONE" or recording.getsampwidth() != 2:
                return True
            rate = recording.getframerate()
            channels = recording.getnchannels()
            samples = array.array("h")
            samples.frombytes(recording.readframes(recording.getnframes()))
    except (wave.Error, EOFError, ValueError):
        return False
    if sys.byteorder != "little":
        samples.byteswap()
    if not samples or rate <= 0 or channels <= 0:
        return False
    window_size = max(channels, (rate // 10) * channels)
    active_seconds = 0.0
    for start in range(0, len(samples), window_size):
        window = samples[start : start + window_size]
        if not window:
            continue
        rms = math.sqrt(sum(sample * sample for sample in window) / len(window))
        peak = max(abs(sample) for sample in window)
        if rms >= 300 and peak >= 800:
            active_seconds += len(window) / (rate * channels)
    return active_seconds >= 0.12


def synthesize_answer_audio(answer: str, api_key: str, language: str) -> bytes:
    speech_style = (
        "Speak clearly in Hindi at a brisk, natural conversational pace, with short pauses."
        if language == "Hindi"
        else "Speak clearly in English at a brisk, natural conversational pace, with short pauses."
    )
    request_body = {
        "contents": [{
            "role": "user",
            "parts": [{"text": answer, "speech_metadata": {"style": speech_style}}],
        }],
        "generationConfig": {
            "responseModalities": ["AUDIO"],
            "speechConfig": {"voiceConfig": {"voice": "Kore"}},
        },
    }
    request = Request(
        "https://generativelanguage.googleapis.com/v1beta/models/gemini-3.8-flash-lite-tts:generateContent",
        data=json.dumps(request_body).encode("utf-8"),
        headers={"Content-Type": "application/json", "x-goog-api-key": api_key},
        method="POST",
    )
    try:
        with urlopen(request, timeout=90) as response:
            payload = json.loads(response.read().decode("utf-8"))
    except HTTPError as error:
        detail = error.read().decode("utf-8", errors="replace")
        try:
            detail = json.loads(detail).get("error", {}).get("message", detail)
        except (AttributeError, json.JSONDecodeError):
            pass
        raise RuntimeError(f"Gemini could not create speech ({error.code}): {detail}") from error
    except URLError as error:
        raise RuntimeError(f"Couldn't reach Gemini to create speech: {error.reason}") from error

    encoded_audio = next(
        (
            part.get("inlineData", {}).get("data")
            for candidate in payload.get("candidates", [])
            for part in candidate.get("content", {}).get("parts", [])
            if part.get("inlineData", {}).get("data")
        ),
        None,
    )
    audio_bytes = base64.b64decode(encoded_audio) if encoded_audio else None
    if not audio_bytes:
        feedback = payload.get("promptFeedback", {}).get("blockReason")
        suffix = f" Gemini returned: {feedback}." if feedback else " Gemini returned no audio."
        raise ValueError("The speech request completed, but no playable voice was returned." + suffix)
    return audio_bytes


def read_upload(upload) -> list[Passage]:
    return read_file(upload.name, upload.getvalue())


def terms(text: str) -> set[str]:
    return {word for word in re.findall(r"[a-zA-Z0-9]+", text.lower()) if word not in STOP_WORDS and len(word) > 1}


def retrieve(question: str, passages: list[Passage], limit: int = 4) -> list[Passage]:
    query_terms = terms(question)
    if not query_terms:
        return []
    ranked: list[tuple[float, int, Passage]] = []
    for index, passage in enumerate(passages):
        text_terms = terms(passage.text)
        overlap = query_terms & text_terms
        if not overlap:
            continue
        # Distinctive terms carry more weight than common words; exact name matches rank well.
        score = sum(2.0 if len(term) >= 7 else 1.0 for term in overlap)
        lowered = passage.text.lower()
        score += sum(1.0 for term in overlap if re.search(rf"\b{re.escape(term)}\b", lowered))
        ranked.append((score, index, passage))
    ranked.sort(key=lambda row: (-row[0], row[1]))
    return [passage for _, _, passage in ranked[:limit]]


def make_tutor_answer(question: str, passages: list[Passage], api_key: str, language: str) -> str:
    sources = "\n\n".join(
        f"[{index}] {passage.filename} — {passage.location}\n{passage.text}"
        for index, passage in enumerate(passages, start=1)
    )
    language_instruction = "Answer in Hindi." if language == "Hindi" else "Answer in English."
    client = genai.Client(api_key=api_key)
    try:
        response = client.models.generate_content(
            model="gemini-3.5-flash-lite",
            contents=(
                "You are VidyaPath, a warm and clear study tutor. Answer the learner using only the provided course excerpts. "
                "Cite factual claims with the exact source labels such as [1] or [2]. If the excerpts do not contain the answer, "
                "say so plainly and ask what material they can add. Never invent facts or citations. Explain simply, then give a short takeaway. "
                f"{language_instruction}\n\nLearner question: {question}\n\nCourse excerpts:\n{sources}"
            ),
        )
    finally:
        client.close()
    return response.text or "I couldn't create an answer from those excerpts. Try asking in a different way."


def generate_quiz(topic: str, passages: list[Passage], api_key: str, language: str, difficulty: str, previous_questions: list[str]) -> list[dict]:
    sources = "\n\n".join(
        f"[{index}] {passage.filename} — {passage.location}\n{passage.text}"
        for index, passage in enumerate(passages, start=1)
    )
    language_instruction = "Write the quiz in Hindi." if language == "Hindi" else "Write the quiz in English."
    previous = "\n".join(f"- {question}" for question in previous_questions) or "None yet."
    prompt = f"""You are VidyaPath, a careful course tutor. Create exactly 5 multiple-choice questions about {topic!r}.
Difficulty: {difficulty}. Make each question answerable from the course excerpts below. Do not use outside knowledge.
{language_instruction}
Do not repeat or lightly reword any question in the excluded list. Use different facts and angles from the excerpts. Within this set, test five distinct details; vary the question forms (for example: who, when, what, why, or how) where the material allows.
Previously asked questions to avoid:
{previous}
Return only valid JSON in this exact shape:
{{"questions":[{{"question":"...","options":{{"A":"...","B":"...","C":"...","D":"..."}},"correct":"A","explanation":"...","source":1}}]}}
For each question, source must be the number of the excerpt that supports its correct answer. Use one of the provided excerpt numbers. Keep explanations concise.

Course excerpts:
{sources}"""
    client = genai.Client(api_key=api_key)
    response = client.models.generate_content(model="gemini-3.5-flash-lite", contents=prompt)
    raw = (response.text or "").strip()
    raw = re.sub(r"^```(?:json)?\s*|\s*```$", "", raw, flags=re.IGNORECASE)
    try:
        payload = json.loads(raw)
    except json.JSONDecodeError:
        start, end = raw.find("{"), raw.rfind("}")
        if start < 0 or end <= start:
            raise ValueError("The tutor returned an unreadable quiz. Please try again.")
        payload = json.loads(raw[start : end + 1])
    questions = payload.get("questions", [])
    clean_questions = []
    for item in questions:
        options = item.get("options", {})
        correct = str(item.get("correct", "")).upper()
        source_number = int(item.get("source", 0))
        if item.get("question") and all(letter in options for letter in ("A", "B", "C", "D")) and correct in options and 1 <= source_number <= len(passages):
            clean_questions.append({
                "question": str(item["question"]),
                "options": {letter: str(options[letter]) for letter in ("A", "B", "C", "D")},
                "correct": correct,
                "explanation": str(item.get("explanation", "")),
                "source": source_number,
            })
    if not clean_questions:
        raise ValueError("No usable questions came back. Try a more specific topic.")
    return clean_questions[:5]


def generate_course_maps(passages: list[Passage], api_key: str) -> list[dict]:
    """Extract source headings, then group related course files into maps."""
    passages_by_file: dict[str, list[int]] = {}
    for index, passage in enumerate(passages):
        passages_by_file.setdefault(passage.filename, []).append(index)
    if not passages_by_file:
        raise ValueError("Upload course material before building a map.")
    if len(passages_by_file) > 12:
        raise ValueError("Please build the map with 12 or fewer files at a time.")

    def read_json(response_text: str) -> dict:
        raw = re.sub(r"^```(?:json)?\s*|\s*```$", "", (response_text or "").strip(), flags=re.IGNORECASE)
        try:
            payload = json.loads(raw)
        except json.JSONDecodeError:
            start_json, end_json = raw.find("{"), raw.rfind("}")
            if start_json < 0 or end_json <= start_json:
                raise ValueError("The course map response could not be read. Please try again.")
            payload = json.loads(raw[start_json : end_json + 1])
        if isinstance(payload, list):
            payload = payload[0] if payload and isinstance(payload[0], dict) else {"topics": payload}
        return payload if isinstance(payload, dict) else {}

    def ask_gemini(prompt: str) -> str:
        client = genai.Client(api_key=api_key)
        try:
            response = client.models.generate_content(model="gemini-3.5-flash-lite", contents=prompt)
            return response.text or ""
        finally:
            client.close()

    def normalized(value: str) -> str:
        return re.sub(r"[^a-z0-9]+", " ", value.casefold()).strip()

    jobs = []
    for filename, indexes in passages_by_file.items():
        for offset in range(0, len(indexes), 16):
            batch_indexes = indexes[offset : offset + 16]
            batch_text = "\n\n".join(
                f"[passage {local_index}] {passages[global_index].location}\n{passages[global_index].text}"
                for local_index, global_index in enumerate(batch_indexes, start=1)
            )
            prompt = f"""Create a clean course map from {filename!r}. Include only main headings and subheadings that appear as their own line in the source, such as numbered section headings, bold headings, or slide titles.
Leave out paragraph text, definitions, examples, explanations, and detail lists. Copy heading titles exactly; do not paraphrase or invent a heading. Set parent to the exact broader heading only when the source clearly shows that relationship; otherwise use null.
Return only JSON: {{"topics":[{{"title":"heading","parent":null,"source":1}}]}}. source is the local passage number shown below.

File passages:\n{batch_text}"""
            jobs.append((filename, batch_indexes, prompt))

    def extract_batch(job) -> tuple[str, list[dict]]:
        filename, batch_indexes, prompt = job
        payload = read_json(ask_gemini(prompt))
        items = payload.get("topics", [])
        if not isinstance(items, list):
            return filename, []
        topics = []
        seen = set()
        for item in items:
            if not isinstance(item, dict):
                continue
            title = str(item.get("title", "")).strip()
            if not title:
                continue
            try:
                local_source = int(item.get("source", 0))
            except (TypeError, ValueError):
                continue
            if not 1 <= local_source <= len(batch_indexes):
                continue
            global_index = batch_indexes[local_source - 1]
            source_lines = [line.strip() for line in passages[global_index].text.splitlines() if line.strip()]
            heading_blocks = {
                normalized(" ".join(source_lines[start : start + count]))
                for count in (1, 2, 3)
                for start in range(max(0, len(source_lines) - count + 1))
            }
            if normalized(title) not in heading_blocks:
                continue
            identity = title.casefold()
            if identity in seen:
                continue
            seen.add(identity)
            topics.append({
                "title": title[:120],
                "parent": str(item.get("parent") or "").strip()[:120],
                "source": global_index + 1,
            })
        return filename, topics

    file_topics: dict[str, list[dict]] = {filename: [] for filename in passages_by_file}
    with ThreadPoolExecutor(max_workers=min(4, len(jobs))) as executor:
        for filename, topics in executor.map(extract_batch, jobs):
            file_topics[filename].extend(topics)

    filenames = list(passages_by_file)
    groups = []
    if len(filenames) == 1:
        groups = [{"course": os.path.splitext(filenames[0])[0], "files": filenames}]
    else:
        outline = {filename: [topic["title"] for topic in file_topics[filename]] for filename in filenames}
        grouping_prompt = f"""Group uploaded course files by subject using their topic titles. Put files in the same map only when their subjects are the same or closely related. Keep completely different subjects in separate maps. Every exact filename must appear once; do not omit files. Give each map a concise course title.
Return only JSON: {{"maps":[{{"course":"...","files":["exact filename"]}}]}}

Course outlines:\n{json.dumps(outline, ensure_ascii=False)}"""
        try:
            grouped = read_json(ask_gemini(grouping_prompt)).get("maps", [])
        except Exception:
            grouped = []
        filename_lookup = {filename.casefold(): filename for filename in filenames}
        assigned = set()
        for group in grouped if isinstance(grouped, list) else []:
            if not isinstance(group, dict):
                continue
            candidates = group.get("files", [])
            if isinstance(candidates, str):
                candidates = [candidates]
            actual_files = []
            for candidate in candidates if isinstance(candidates, list) else []:
                actual = filename_lookup.get(str(candidate).strip().casefold())
                if actual and actual not in assigned:
                    actual_files.append(actual)
                    assigned.add(actual)
            if actual_files:
                groups.append({
                    "course": str(group.get("course") or os.path.splitext(actual_files[0])[0]).strip()[:90],
                    "files": actual_files,
                })
        for filename in filenames:
            if filename not in assigned:
                groups.append({"course": os.path.splitext(filename)[0], "files": [filename]})

    course_maps = []
    for group in groups:
        topics = []
        seen_titles = set()
        for filename in group["files"]:
            for topic in file_topics[filename]:
                title_key = topic["title"].casefold()
                if title_key in seen_titles:
                    continue
                seen_titles.add(title_key)
                topics.append(dict(topic))
        if not topics:
            continue

        title_lookup = {topic["title"].casefold(): topic for topic in topics}
        course_title = group["course"]
        for topic in topics:
            parent_key = topic["parent"].casefold()
            if parent_key not in title_lookup or parent_key == topic["title"].casefold():
                topic["parent"] = course_title
            else:
                topic["parent"] = title_lookup[parent_key]["title"]
        for topic in topics:
            visited = {topic["title"].casefold()}
            parent = topic["parent"]
            while parent.casefold() in title_lookup:
                parent_key = parent.casefold()
                if parent_key in visited:
                    topic["parent"] = course_title
                    break
                visited.add(parent_key)
                parent = title_lookup[parent_key]["parent"]
        course_maps.append({"course": course_title, "files": group["files"], "topics": topics})

    if not course_maps:
        raise ValueError("No clear headings were found in the uploaded materials. Try text-based course files with visible headings.")
    return course_maps

def course_map_graph(course_map: dict) -> str:
    """Build a polished, readable left-to-right course tree."""
    def quote(value: str) -> str:
        return json.dumps(value, ensure_ascii=False)

    topics = course_map["topics"]
    title_to_id = {item["title"].casefold(): f"topic_{index}" for index, item in enumerate(topics)}
    course_label = "\n".join(textwrap.wrap(course_map["course"], width=24, break_long_words=False))
    lines = [
        "digraph CourseMap {",
        '  graph [rankdir=LR, bgcolor="transparent", pad="0.35", nodesep="0.55", ranksep="0.9", splines=curved, ordering=out];',
        '  node [shape="box", style="rounded,filled", fontname="Arial", fontsize=17, fontcolor="#192944", color="#d8e2f2", penwidth=1.8, margin="0.28,0.22", width=2.1, height=0.72];',
        '  edge [color="#8295b6", penwidth=2.4, arrowsize=0.95, arrowhead=vee];',
        f'  course_root [label={quote(course_label)}, fillcolor="#3347a8", fontcolor="#ffffff", color="#3347a8", fontsize=20, penwidth=0, margin="0.36,0.28"];',
        '  graph [rankdir=LR, bgcolor="transparent", pad="0.2", nodesep="0.28", ranksep="0.48", splines=curved, ordering=out];',
        '  node [shape="box", style="rounded,filled", fontname="Arial", fontsize=25, fontcolor="#192944", color="#d8e2f2", penwidth=1.6, margin="0.2,0.15", width=1.8, height=0.62];',
        '  edge [color="#8295b6", penwidth=1.8, arrowsize=0.75, arrowhead=vee];',
        f'  course_root [label={quote(course_label)}, fillcolor="#3347a8", fontcolor="#ffffff", color="#3347a8", fontsize=27, penwidth=0, margin="0.25,0.2"];',
    ]
    palette = ["#edf1ff", "#e7f7f3", "#fff3df", "#f3edff"]
    for index, item in enumerate(topics):
        label = "\n".join(textwrap.wrap(item["title"], width=22, break_long_words=False, break_on_hyphens=False))
        lines.append(f"  topic_{index} [label={quote(label)}, fillcolor={quote(palette[index % len(palette)])}];")
        label = "\n".join(textwrap.wrap(item["title"], width=26, break_long_words=False, break_on_hyphens=False))
        font_size = max(21, 29 - max(0, len(item["title"]) - 16) // 3)
        lines.append(f"  topic_{index} [label={quote(label)}, fillcolor={quote(palette[index % len(palette)])}, fontsize={font_size}];")
    for item in topics:
        target = title_to_id[item["title"].casefold()]
        source = title_to_id.get(item["parent"].casefold(), "course_root")
        lines.append(f"  {source} -> {target};")
    lines.append("}")
    return "\n".join(lines)


def get_gemini_api_key() -> str:
    return get_secret("GEMINI_API_KEY")


st.set_page_config(page_title="VidyaPath · AI Study Companion", page_icon="✦", layout="wide")
st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700&family=Manrope:wght@500;600;700;800&display=swap');
    :root { --ink:#17253d; --muted:#70809a; --blue:#4361ee; --teal:#13a89e; --paper:#f6f8fc; }
    .stApp { background:radial-gradient(ellipse at 8% 8%,#e5e8ff 0,transparent 34%),radial-gradient(ellipse at 92% 78%,#d8f3f0 0,transparent 32%),linear-gradient(135deg,#f7f8ff 0%,#eef2fb 52%,#f7f9ff 100%); color:var(--ink); font-family:'DM Sans',sans-serif; }
    [data-testid="stSidebar"] { background: #13213a; }
    [data-testid="stSidebar"] * { color:#eef3ff; }
    [data-testid="stSidebar"] [data-testid="stFileUploader"] section { background:#1b2c49; border:1px dashed #6680a9; }
    h1,h2,h3 { font-family:'Manrope',sans-serif !important; color:var(--ink); font-weight:800 !important; }
    [data-testid="stTabs"] [role="tab"] { font-weight:700 !important; }
    [data-testid="stSidebar"] h1,[data-testid="stSidebar"] h2,[data-testid="stSidebar"] h3 { color:#fff !important; }
    .block-container { max-width:1250px; padding-top:2.2rem; padding-bottom:3rem; }
    .hero { padding:2rem 2.2rem; border-radius:24px; color:white; background:linear-gradient(115deg,#243b83 0%,#4361ee 56%,#13a89e 120%); box-shadow:0 18px 40px #31498c26; margin-bottom:1.8rem; }
    .hero h1 { color:white !important; font-size:2.6rem; margin:0 0 .35rem 0; }
    .hero p { color:#e8edff; margin:0; font-size:1.05rem; }
    .profile-avatar-row { display:flex; justify-content:center; align-items:center; min-height:3.5rem; margin:.15rem 0 .45rem; }
    .profile-avatar { width:48px; height:48px; border-radius:50%; display:flex; align-items:center; justify-content:center; background:linear-gradient(135deg,#4361ee,#13a89e); color:#fff; border:2px solid #fff; box-shadow:0 7px 18px #233b6430; font:800 1.15rem 'Manrope',sans-serif; }
    .brand-lockup { display:flex; align-items:center; gap:.72rem; margin:.2rem 0 .35rem; }
    .brand-lockup svg { width:42px; height:42px; flex:0 0 42px; filter:drop-shadow(0 5px 10px #07152d55); }
    .brand-lockup span { color:#fff; font:800 1.55rem 'Manrope',sans-serif; letter-spacing:-.035em; }
    .eyebrow { text-transform:uppercase; letter-spacing:.16em; font-size:.72rem; font-weight:700; opacity:.8; margin-bottom:.45rem; }
    .card { background:white; border:1px solid #e8edf5; border-radius:18px; padding:1.25rem 1.35rem; box-shadow:0 8px 24px #233b6410; margin-bottom:1rem; }
    .metric { color:var(--muted); font-size:.86rem; }
    .metric strong { color:var(--ink); display:block; font:700 1.45rem 'Manrope',sans-serif; margin-top:.25rem; }
    [data-testid="stMetric"] { background:#ffffff !important; border:1px solid #dce5f2 !important; border-radius:16px !important; padding:1rem 1.15rem !important; box-shadow:0 8px 24px #233b6410 !important; }
    [data-testid="stMetricLabel"], [data-testid="stMetricLabel"] *, [data-testid="stMetricValue"], [data-testid="stMetricValue"] *, [data-testid="stMetricDelta"], [data-testid="stMetricDelta"] * { color:#17253d !important; opacity:1 !important; }
    [data-testid="stMetricLabel"] { font-weight:600 !important; }
    .stButton > button[kind="primary"], .stButton > button[data-testid="stBaseButton-primary"] { background:linear-gradient(110deg,#3457d5 0%,#6541d9 100%) !important; border:0 !important; border-radius:12px !important; font-weight:700 !important; min-height:2.8rem; box-shadow:0 7px 20px #6541d940 !important; transition:transform .16s ease, box-shadow .16s ease, filter .16s ease; }
    [data-testid="stFormSubmitButton"] button { background:linear-gradient(110deg,#3457d5 0%,#6541d9 100%) !important; color:#fff !important; border:0 !important; border-radius:12px !important; font-weight:700 !important; min-height:2.8rem; box-shadow:0 7px 20px #6541d940 !important; transition:transform .16s ease, box-shadow .16s ease, filter .16s ease; }
    [data-testid="stFormSubmitButton"] button:hover { filter:brightness(1.06); transform:translateY(-2px); box-shadow:0 11px 24px #6541d955 !important; }
    [data-testid="stFormSubmitButton"] button:active { transform:translateY(0); box-shadow:0 4px 10px #6541d935 !important; }
    .stButton > button[kind="primary"], .stButton > button[kind="primary"] *, .stButton > button[data-testid="stBaseButton-primary"], .stButton > button[data-testid="stBaseButton-primary"] * { color:#fff !important; }
    .stButton > button:not([kind="primary"]) { background:linear-gradient(105deg,#4361ee,#13a89e) !important; color:#fff !important; border:0 !important; border-radius:12px; font-weight:700; box-shadow:0 7px 18px #4361ee40; transition:transform .16s ease, box-shadow .16s ease, filter .16s ease; }
    .stButton > button:not([kind="primary"]) * { color:#fff !important; }
    .stButton > button[kind="primary"]:hover, .stButton > button[data-testid="stBaseButton-primary"]:hover, .stButton > button:not([kind="primary"]):hover { filter:brightness(1.06); transform:translateY(-2px); box-shadow:0 11px 24px #6541d955; }
    .stButton > button[kind="primary"]:active, .stButton > button[data-testid="stBaseButton-primary"]:active, .stButton > button:not([kind="primary"]):active { transform:translateY(0); box-shadow:0 4px 10px #6541d935; }
    .stTextInput input { border-radius:11px; background:#fff !important; color:#17253d !important; border:1px solid #d5dfef !important; min-height:3rem; }
    .stTextInput input::placeholder { color:#8491a7 !important; opacity:1 !important; }
    .stTextInput input:focus { border-color:#6278ee !important; box-shadow:0 0 0 3px #4361ee20 !important; }
    .stSelectbox [data-baseweb="select"] { border-radius:11px; }
    .st-key-study_note_source label, .st-key-study_note_text label,
    .st-key-study_note_document_file label, .st-key-study_note_document_location label,
    .st-key-study_note_source label *, .st-key-study_note_text label *,
    .st-key-study_note_document_file label *, .st-key-study_note_document_location label * { color:#17253d !important; opacity:1 !important; font-weight:600 !important; }
    .st-key-study_note_source [data-baseweb="select"] > div,
    .st-key-study_note_document_file [data-baseweb="select"] > div,
    .st-key-study_note_document_location [data-baseweb="select"] > div { background:#fff !important; color:#17253d !important; border:1px solid #bdcbe0 !important; border-radius:11px !important; }
    .st-key-study_note_source [data-baseweb="select"] input,
    .st-key-study_note_document_file [data-baseweb="select"] input,
    .st-key-study_note_document_location [data-baseweb="select"] input { color:#17253d !important; caret-color:#3457d5 !important; }
    .st-key-study_note_source [data-baseweb="select"] svg,
    .st-key-study_note_document_file [data-baseweb="select"] svg,
    .st-key-study_note_document_location [data-baseweb="select"] svg { fill:#52627a !important; }
    [data-testid="stSelectbox"] [data-baseweb="select"] > div { background:#fff !important; background-color:#fff !important; color:#17253d !important; border:1px solid #bdcbe0 !important; border-radius:11px !important; }
    [data-testid="stSelectbox"] [data-baseweb="select"] span,
    [data-testid="stSelectbox"] [data-baseweb="select"] input { color:#17253d !important; -webkit-text-fill-color:#17253d !important; }
    [data-testid="stSelectbox"] [data-baseweb="select"] svg { fill:#52627a !important; color:#52627a !important; }
    .st-key-study_note_text textarea { background:#fff !important; color:#17253d !important; caret-color:#3457d5 !important; border:1px solid #bdcbe0 !important; border-radius:11px !important; min-height:9rem !important; color-scheme:light; }
    .st-key-study_note_text textarea::placeholder { color:#77859b !important; opacity:1 !important; }
    .st-key-study_note_text textarea:focus { border-color:#4361ee !important; box-shadow:0 0 0 3px #4361ee25 !important; outline:none !important; }
    [data-baseweb="popover"] [role="listbox"] { background:#fff !important; }
    [data-baseweb="popover"] [role="option"] { color:#17253d !important; }
    [data-testid="stTextInput"] label, [data-testid="stTextInput"] label *, [data-testid="stTextInput"] p { color:#17253d !important; opacity:1 !important; }
    [data-testid="stLinkButton"] a { background:linear-gradient(105deg,#4361ee,#13a89e) !important; color:#fff !important; border:0 !important; border-radius:12px !important; font-weight:700 !important; box-shadow:0 7px 18px #4361ee40; }
    [data-testid="stLinkButton"] a * { color:#fff !important; }
    [data-testid="stLinkButton"] a:hover { filter:brightness(1.06); transform:translateY(-2px); box-shadow:0 11px 24px #4361ee55; }
    [data-testid="stRadio"] label, [data-testid="stRadio"] label *, [data-testid="stRadio"] p { color:#17253d !important; opacity:1 !important; }
    [data-baseweb="tab"] { color:#52627a !important; }
    [data-baseweb="tab"] *, [role="tab"] { color:#52627a !important; }
    [role="tab"][aria-selected="true"], [role="tab"][aria-selected="true"] * { color:#354fc7 !important; }
    [data-testid="stTabs"] [data-baseweb="tab-list"] { gap:.55rem; border-bottom:1px solid #dce4f0; }
    [data-testid="stTabs"] [data-baseweb="tab"] { background:#e9eef8 !important; border-radius:11px 11px 0 0; padding:.65rem 1rem; transition:background .16s ease, color .16s ease; }
    [data-testid="stTabs"] [data-baseweb="tab"][aria-selected="true"] { background:linear-gradient(110deg,#3457d5,#6541d9) !important; border-bottom-color:#6541d9 !important; }
    [data-testid="stTabs"] [data-baseweb="tab"][aria-selected="true"], [data-testid="stTabs"] [data-baseweb="tab"][aria-selected="true"] * { color:#fff !important; }
    .st-key-google_auth_area .stButton > button { background:linear-gradient(110deg,#14532d 0%,#15803d 58%,#6d28d9 100%) !important; border:1px solid #ffffff35 !important; color:#fff !important; border-radius:13px !important; font-weight:700 !important; min-height:2.9rem; box-shadow:0 8px 22px #4b257944 !important; }
    .st-key-google_auth_area .stButton > button * { color:#fff !important; }
    .st-key-google_auth_area [data-testid="stLinkButton"] a { background:linear-gradient(110deg,#14532d 0%,#15803d 58%,#6d28d9 100%) !important; border:1px solid #ffffff35 !important; color:#fff !important; border-radius:13px !important; box-shadow:0 8px 22px #4b257944 !important; }
    .st-key-google_auth_area [data-testid="stLinkButton"] a * { color:#fff !important; }
    [data-testid="stExpander"] { background:white; border:1px solid #e8edf5; border-radius:12px; }
    .small-note { color:var(--muted); font-size:.82rem; }
    /* Clearer, more polished type and consistent spacing across the learning workspace. */
    .stApp, .stApp button, .stApp input, .stApp textarea, .stApp select { font-family:'DM Sans',sans-serif; }
    [data-testid="stMarkdownContainer"] p, [data-testid="stMarkdownContainer"] li { font-size:1.02rem; line-height:1.65; font-weight:500; }
    [data-testid="stMarkdownContainer"] strong, [data-testid="stMarkdownContainer"] b { font-weight:750 !important; }
    h1 { font-size:2.35rem !important; letter-spacing:-.035em; }
    h2 { font-size:1.8rem !important; letter-spacing:-.025em; margin-top:1.8rem !important; }
    h3 { font-size:1.35rem !important; letter-spacing:-.015em; }
    [data-testid="stWidgetLabel"], [data-testid="stWidgetLabel"] * { font-size:1rem !important; font-weight:700 !important; color:#20304b !important; }
    [data-testid="stSidebar"] [data-testid="stWidgetLabel"], [data-testid="stSidebar"] [data-testid="stWidgetLabel"] * { color:#eef3ff !important; }
    [data-testid="stCaptionContainer"] { font-size:.92rem !important; line-height:1.55 !important; font-weight:500 !important; color:#64748b !important; }
    [data-testid="stSidebar"] [data-testid="stCaptionContainer"] { color:#c1cce0 !important; }
    [data-testid="stMetricLabel"] { font-size:.92rem !important; letter-spacing:.035em; text-transform:uppercase; }
    [data-testid="stMetricValue"] { font-family:'Manrope',sans-serif !important; font-size:1.85rem !important; font-weight:800 !important; }
    .stButton > button, [data-testid="stLinkButton"] a, [data-testid="stFormSubmitButton"] button { font-size:1rem !important; letter-spacing:.005em; }
    [data-testid="stVerticalBlockBorderWrapper"] { border-color:#dbe4f2 !important; border-radius:18px !important; box-shadow:0 8px 24px #233b640b; }
    [data-testid="stTabs"] [data-baseweb="tab"] { font-size:.98rem !important; font-weight:700 !important; }

    /* ---- Readable messages: dark text on a solid tinted background, in light AND dark browser themes ---- */
    [data-testid="stAlert"] { border-radius:14px !important; font-size:1rem !important; line-height:1.55 !important; font-weight:600 !important; padding:.2rem .35rem !important; background:#eef2ff !important; border:1px solid #c7d2fe !important; border-left:6px solid #4361ee !important; box-shadow:0 6px 18px #233b6414; }
    [data-testid="stAlert"], [data-testid="stAlert"] *, [data-testid="stAlert"] p, [data-testid="stAlert"] span, [data-testid="stAlert"] div { color:#17253d !important; -webkit-text-fill-color:#17253d !important; opacity:1 !important; }
    [data-testid="stAlert"] svg { fill:currentColor; }
    [data-testid="stAlert"]:has([data-testid="stAlertContentWarning"]) { background:#fff1c2 !important; border-color:#f5c542 !important; border-left-color:#d97706 !important; }
    [data-testid="stAlert"]:has([data-testid="stAlertContentWarning"]), [data-testid="stAlert"]:has([data-testid="stAlertContentWarning"]) * { color:#5c3a00 !important; -webkit-text-fill-color:#5c3a00 !important; }
    [data-testid="stAlert"]:has([data-testid="stAlertContentError"]) { background:#ffe1e1 !important; border-color:#f5a3a3 !important; border-left-color:#dc2626 !important; }
    [data-testid="stAlert"]:has([data-testid="stAlertContentError"]), [data-testid="stAlert"]:has([data-testid="stAlertContentError"]) * { color:#7f1111 !important; -webkit-text-fill-color:#7f1111 !important; }
    [data-testid="stAlert"]:has([data-testid="stAlertContentSuccess"]) { background:#d9f7e6 !important; border-color:#86d9ab !important; border-left-color:#16a34a !important; }
    [data-testid="stAlert"]:has([data-testid="stAlertContentSuccess"]), [data-testid="stAlert"]:has([data-testid="stAlertContentSuccess"]) * { color:#0b4a24 !important; -webkit-text-fill-color:#0b4a24 !important; }
    [data-testid="stAlert"]:has([data-testid="stAlertContentInfo"]) { background:#e0ecff !important; border-color:#9dbcf7 !important; border-left-color:#2563eb !important; }
    [data-testid="stAlert"]:has([data-testid="stAlertContentInfo"]), [data-testid="stAlert"]:has([data-testid="stAlertContentInfo"]) * { color:#12306b !important; -webkit-text-fill-color:#12306b !important; }
    [data-testid="stAlert"] > div, [data-testid="stAlert"] [data-testid^="stAlertContent"] { background:transparent !important; border:0 !important; box-shadow:none !important; }
    ::selection { background:#c7d2fe; color:#17253d; }
    /* Sidebar controls with a light background need dark text */
    [data-testid="stSidebar"] [data-testid="stFileUploader"] button, [data-testid="stSidebar"] [data-testid="stFileUploader"] button * { background:#fff !important; color:#243b83 !important; -webkit-text-fill-color:#243b83 !important; font-weight:700 !important; }
    [data-testid="stSidebar"] [data-baseweb="select"] > div, [data-testid="stSidebar"] [data-baseweb="select"] * { color:#17253d !important; -webkit-text-fill-color:#17253d !important; }
    [data-testid="stSidebar"] [data-baseweb="select"] svg { fill:#52627a !important; }
    [data-testid="stSidebar"] [data-testid="stSelectbox"] div, [data-testid="stSidebar"] [data-testid="stSelectbox"] input, [data-testid="stSidebar"] [data-testid="stSelectbox"] svg { color:#17253d !important; -webkit-text-fill-color:#17253d !important; fill:#52627a !important; }
    [data-testid="stSidebar"] [data-testid="stSelectbox"] [data-testid="stWidgetLabel"], [data-testid="stSidebar"] [data-testid="stSelectbox"] [data-testid="stWidgetLabel"] * { color:#eef3ff !important; -webkit-text-fill-color:#eef3ff !important; }

    /* ---- Extra polish: lively but calm ---- */
    @keyframes vp-rise { from { opacity:0; transform:translateY(10px); } to { opacity:1; transform:none; } }
    @keyframes vp-glow { 0%,100% { transform:translate(0,0) scale(1); } 50% { transform:translate(-18px,12px) scale(1.08); } }
    .block-container > div { animation:vp-rise .45s ease both; }
    [data-testid="stSidebar"] { background:linear-gradient(180deg,#14223d 0%,#1b2f5c 55%,#183f5a 100%) !important; }
    .hero { position:relative; overflow:hidden; }
    .hero:before { content:""; position:absolute; width:300px; height:300px; right:-70px; top:-130px; border-radius:50%; background:radial-gradient(circle,#ffffff3a 0%,#ffffff00 70%); animation:vp-glow 9s ease-in-out infinite; pointer-events:none; }
    .hero h1 { text-shadow:0 2px 14px #0d1b4a40; }
    .card { transition:transform .18s ease, box-shadow .18s ease, border-color .18s ease; }
    .card:hover { transform:translateY(-3px); box-shadow:0 16px 34px #233b6422; border-color:#c5d3ee; }
    .metric { position:relative; overflow:hidden; }
    .metric:before { content:""; position:absolute; left:0; top:0; bottom:0; width:5px; background:linear-gradient(180deg,#4361ee,#13a89e); }
    .metric strong { background:linear-gradient(110deg,#243b83,#13a89e); -webkit-background-clip:text; background-clip:text; color:transparent; }
    .path-step { display:flex; align-items:center; gap:.85rem; padding:.7rem .5rem; border-radius:14px; transition:background .16s ease, transform .16s ease; }
    .path-step:hover { background:#f1f5ff; transform:translateX(4px); }
    .path-icon { flex:0 0 40px; height:40px; border-radius:12px; display:flex; align-items:center; justify-content:center; color:#fff; font-size:1.1rem; font-weight:800; box-shadow:0 6px 14px #233b6428; }
    .path-step b { display:block; color:var(--ink); font:700 1rem 'Manrope',sans-serif; }
    .path-step small { color:var(--muted); font-size:.84rem; }
    [data-testid="stTabs"] [data-baseweb="tab"]:hover { background:#dbe4fb !important; }
    .stTextInput input:hover, [data-testid="stTextArea"] textarea:hover { border-color:#9fb1e8 !important; }
    [data-testid="stExpander"] { box-shadow:0 6px 18px #233b640d; }
    [data-testid="stVerticalBlockBorderWrapper"] { background:#ffffffb8; backdrop-filter:blur(6px); transition:box-shadow .18s ease, transform .18s ease; }
    [data-testid="stVerticalBlockBorderWrapper"]:hover { box-shadow:0 14px 30px #233b6418; }
    ::-webkit-scrollbar { width:10px; height:10px; } ::-webkit-scrollbar-thumb { background:#b9c6e6; border-radius:8px; }
    /* Section selector styled like tabs */
    .st-key-active_section [role="radiogroup"] { gap:.3rem !important; flex-wrap:wrap; border-bottom:1px solid #dce4f0; padding-bottom:.45rem; margin-bottom:.6rem; }
    .st-key-active_section label { background:#e9eef8; border-radius:11px; padding:.5rem .62rem !important; margin:0 !important; cursor:pointer; transition:background .16s ease, transform .16s ease; }
    .st-key-active_section label:hover { background:#dbe4fb; transform:translateY(-1px); }
    .st-key-active_section label > div > div:not([data-testid="stMarkdownContainer"]) { display:none !important; }
    .st-key-active_section label p { margin:0 !important; font-weight:700 !important; font-size:.9rem !important; white-space:nowrap; color:#52627a !important; }
    .st-key-active_section label[data-selected="true"], .st-key-active_section label:has(input:checked) { background:linear-gradient(110deg,#3457d5,#6541d9); box-shadow:0 7px 18px #6541d940; }
    .st-key-active_section label[data-selected="true"] *, .st-key-active_section label:has(input:checked) * { color:#fff !important; -webkit-text-fill-color:#fff !important; }
    @media (prefers-reduced-motion: reduce) { .block-container > div, .hero:before { animation:none !important; } }
    @media (max-width: 760px) {
        .block-container { padding-left:1rem !important; padding-right:1rem !important; }
        .hero { padding:1.5rem !important; border-radius:20px !important; }
        .hero h1 { font-size:2rem !important; }
        h2 { font-size:1.55rem !important; }
        [data-testid="stMarkdownContainer"] p, [data-testid="stMarkdownContainer"] li { font-size:1rem; }
    }
    </style>
    """,
    unsafe_allow_html=True,
)

auth_user = st.session_state.get("auth_user")
oauth_code = st.query_params.get("code")
password_reset_redirect = st.query_params.get("reset_password") == "1"
supabase = None
if auth_user or st.session_state.get("show_login", False) or oauth_code:
    supabase_url = get_secret("SUPABASE_URL")
    supabase_anon_key = get_secret("SUPABASE_ANON_KEY")
    if not supabase_url or not supabase_anon_key:
        st.error("Sign-in is not connected yet. Add the Supabase project settings to the app's private configuration.")
        if st.button("← Back to your workspace"):
            st.session_state.show_login = False
            st.rerun()
        st.stop()
    try:
        supabase = get_supabase_client()
    except Exception:
        st.error("VidyaPath couldn't connect to its authentication service. Check the private Supabase settings and try again.")
        if st.button("← Back to your workspace"):
            st.session_state.show_login = False
            st.rerun()
        st.stop()

    if oauth_code and not auth_user:
        try:
            supabase.auth.exchange_code_for_session({"auth_code": oauth_code})
            user_response = supabase.auth.get_user()
            auth_user = user_response.user
            st.session_state.auth_user = {"id": str(auth_user.id), "email": auth_user.email or ""}
            st.session_state.show_login = False
            st.session_state.course_file_uploader_version = st.session_state.get("course_file_uploader_version", 0) + 1
            if password_reset_redirect:
                st.session_state.password_recovery = True
            st.session_state.pop("google_auth_url", None)
            st.query_params.clear()
            st.rerun()
        except Exception:
            st.query_params.clear()
            st.session_state.show_login = True
            st.error("Google sign-in could not be completed. Please try again.")

if st.session_state.get("password_recovery"):
    st.markdown("<div class='hero'><div class='eyebrow'>Account security</div><h1>Choose a new password</h1><p>Set a fresh password for your VidyaPath account.</p></div>", unsafe_allow_html=True)
    with st.form("set_new_password"):
        recovery_password = st.text_input("New password", type="password", placeholder="New password (8+ characters)", label_visibility="collapsed")
        recovery_password_confirm = st.text_input("Confirm new password", type="password", placeholder="Confirm your new password", label_visibility="collapsed")
        recovery_submitted = st.form_submit_button("Update password", type="primary", use_container_width=True)
    if recovery_submitted:
        if len(recovery_password) < 8:
            st.warning("Choose a password with at least 8 characters.")
        elif recovery_password != recovery_password_confirm:
            st.warning("The two passwords don't match.")
        else:
            try:
                supabase.auth.update_user({"password": recovery_password})
                supabase.auth.sign_out()
                st.session_state.pop("auth_user", None)
                st.session_state.pop("password_recovery", None)
                st.session_state.password_reset_success = True
                st.session_state.show_login = True
                st.rerun()
            except Exception:
                st.error("VidyaPath couldn't update the password. Request a new reset email and try again.")
    st.stop()

if not st.session_state.get("auth_user") and st.session_state.get("show_login", False):
    st.markdown(
        """
        <style>
        [data-testid="stAppViewContainer"] { background:radial-gradient(ellipse at 12% 8%,#dedbff 0,transparent 33%),radial-gradient(ellipse at 90% 84%,#c9f1e8 0,transparent 32%),linear-gradient(135deg,#f8f7ff 0%,#edf1ff 52%,#f5faff 100%) !important; }
        [data-testid="stHeader"] { background:transparent !important; }
        .block-container { max-width:1040px !important; padding-top:3.1rem !important; }
        .hero { position:relative; overflow:hidden; padding:3rem 3.2rem !important; border:1px solid #ffffff65; background:linear-gradient(120deg,#21194b 0%,#4936b0 52%,#137f93 112%) !important; box-shadow:0 26px 70px #352b8050 !important; }
        .hero:after { content:""; position:absolute; width:250px; height:250px; top:-145px; right:-35px; border-radius:50%; background:radial-gradient(circle,#ffffff35 0%,#ffffff00 72%); pointer-events:none; }
        .hero h1 { font-size:2.9rem !important; letter-spacing:-.045em; }
        [data-testid="stTabs"] { background:#ffffffc9; border:1px solid #ffffff; border-radius:22px; padding:1.25rem 1.4rem 1.5rem; box-shadow:0 22px 55px #302c7030; backdrop-filter:blur(16px); }
        [data-testid="stTabs"] [data-baseweb="tab-list"] { background:transparent; }
        [data-testid="stTabs"] [data-baseweb="tab"] { border-radius:12px 12px 0 0; }
        @media (max-width:700px) { .block-container { padding:1.25rem 1rem 2rem !important; } .hero { padding:2rem 1.5rem !important; } .hero h1 { font-size:2.2rem !important; } }
        </style>
        """,
        unsafe_allow_html=True,
    )
    if st.session_state.pop("password_reset_success", False):
        st.success("Your password has been changed. Sign in with your new password.")
    back_left, back_space = st.columns([1, 8])
    with back_left:
        if st.button("← Home", key="login_back_home"):
            st.session_state.show_login = False
            st.session_state.pop("google_auth_url", None)
            st.rerun()
    st.markdown("<div class='hero'><div class='eyebrow'>Your private learning space</div><h1>Welcome to VidyaPath.</h1><p>Sign in to keep your course files and learning progress with your account.</p></div>", unsafe_allow_html=True)
    sign_in_tab, create_account_tab = st.tabs(["Sign in", "Create account"])
    with sign_in_tab:
        with st.form("email_sign_in"):
            sign_in_email = st.text_input("Email address", placeholder="Enter your email address", label_visibility="collapsed")
            sign_in_password = st.text_input("Password", type="password", placeholder="Enter your password", label_visibility="collapsed")
            sign_in_submitted = st.form_submit_button("Sign in", type="primary", use_container_width=True)
        if sign_in_submitted:
            try:
                auth_response = supabase.auth.sign_in_with_password({"email": sign_in_email.strip(), "password": sign_in_password})
                st.session_state.auth_user = {"id": str(auth_response.user.id), "email": auth_response.user.email or sign_in_email.strip()}
                st.session_state.show_login = False
                st.session_state.course_file_uploader_version = st.session_state.get("course_file_uploader_version", 0) + 1
                st.rerun()
            except Exception:
                # Keep provider details private and give the same clear message
                # for incorrect credentials or an account blocked by provider policy.
                st.error("We couldn't sign you in. Check your email and password, then try again.")
        with st.expander("Forgot password?"):
            with st.form("request_password_reset"):
                reset_email = st.text_input("Account email", placeholder="Enter your account email", label_visibility="collapsed", key="reset_email")
                reset_requested = st.form_submit_button("Send password reset email", use_container_width=True)
            if reset_requested:
                if not reset_email.strip():
                    st.warning("Enter the email address for your account.")
                else:
                    try:
                        app_url = get_secret("APP_URL", "http://localhost:8501").rstrip("/")
                        supabase.auth.reset_password_for_email(
                            reset_email.strip(),
                            {"redirect_to": f"{app_url}/?reset_password=1"},
                        )
                        st.success("If that account can receive a reset link, an email will arrive shortly. Check your inbox and Spam/Junk.")
                    except Exception:
                        st.error("Couldn't send the reset email. Check Supabase Auth email settings and the allowed redirect URL.")
    with create_account_tab:
        with st.form("email_sign_up"):
            new_email = st.text_input("Email address", placeholder="Enter your email address", label_visibility="collapsed", key="new_email")
            new_password = st.text_input("Create a password", type="password", placeholder="Create a password (8+ characters)", label_visibility="collapsed", key="new_password", help="Use at least 8 characters.")
            sign_up_submitted = st.form_submit_button("Create account", type="primary", use_container_width=True)
        if sign_up_submitted:
            if len(new_password) < 8:
                st.warning("Choose a password with at least 8 characters.")
            else:
                try:
                    auth_response = supabase.auth.sign_up({"email": new_email.strip(), "password": new_password})
                    if auth_response.session and auth_response.user:
                        st.session_state.auth_user = {"id": str(auth_response.user.id), "email": auth_response.user.email or new_email.strip()}
                        st.session_state.show_login = False
                        st.session_state.course_file_uploader_version = st.session_state.get("course_file_uploader_version", 0) + 1
                        st.rerun()
                    st.info("Account created. If sign-in is still blocked, ask the app owner to turn off email confirmation in Supabase Auth.")
                except Exception:
                    st.error("We couldn't create that account. Check the email address and try again.")

    st.markdown("<div style='text-align:center; margin:1.2rem 0; color:#70809a;'>or</div>", unsafe_allow_html=True)
    with st.container(key="google_auth_area"):
        if st.session_state.get("google_auth_url"):
            st.link_button("✦ Continue with Google", st.session_state.google_auth_url, use_container_width=True)
        elif st.button("✦ Continue with Google", use_container_width=True):
            try:
                oauth_response = supabase.auth.sign_in_with_oauth({
                    "provider": "google",
                    "options": {"redirect_to": get_secret("APP_URL", "http://localhost:8501")},
                })
                st.session_state.google_auth_url = oauth_response.url
                st.rerun()
            except Exception:
                st.error("Google sign-in isn't configured yet. The app owner needs to enable Google in Supabase first.")
    st.stop()

is_authenticated = bool(st.session_state.get("auth_user"))
user_id = st.session_state.auth_user["id"] if is_authenticated else None
uploads = None

def sign_out_current_user():
    try:
        supabase.auth.sign_out()
    finally:
        st.session_state["voice_recording_version"] = st.session_state.get("voice_recording_version", 0) + 1
        for key in ("auth_user", "active_user_id", "passages", "loaded_files", "stored_file_paths", "file_paths_by_name", "removed_upload_paths", "files_needing_ocr", "video_sources", "video_notice", "voice_result", "answers", "tutor_history", "tutor_history_error", "study_notes", "study_notes_error", "quiz_history", "current_quiz", "course_map", "course_maps", "google_auth_url"):
            st.session_state.pop(key, None)
        st.session_state.pop("supabase_client", None)
        st.session_state.course_file_uploader_version = st.session_state.get("course_file_uploader_version", 0) + 1
        st.session_state.show_login = False
        st.rerun()

with st.sidebar:
    st.markdown(
        "<div class='brand-lockup'><svg viewBox='0 0 48 48' role='img' aria-label='VidyaPath logo'>"
        "<defs><linearGradient id='vidyapathLogoGradient' x1='0' y1='0' x2='1' y2='1'>"
        "<stop offset='0%' stop-color='#5574ff'/><stop offset='100%' stop-color='#12b6ae'/></linearGradient></defs>"
        "<rect x='2' y='2' width='44' height='44' rx='14' fill='url(#vidyapathLogoGradient)'/>"
        "<path d='M9 14.5c5.5-1.5 10.4-.2 15 3.4v17c-4.6-3.2-9.5-4.3-15-2.6z "
        "M39 14.5c-5.5-1.5-10.4-.2-15 3.4v17c4.6-3.2 9.5-4.3 15-2.6z' "
        "fill='none' stroke='white' stroke-width='2.5' stroke-linecap='round' stroke-linejoin='round'/>"
        "<path d='M13 20c3.1-.6 5.8 0 8.2 1.6M13 25c3.1-.6 5.8 0 8.2 1.6M35 20c-3.1-.6-5.8 0-8.2 1.6M35 25c-3.1-.6-5.8 0-8.2 1.6' "
        "fill='none' stroke='#b9fff0' stroke-width='1.6' stroke-linecap='round'/>"
        "<path d='M24 18v17' stroke='white' stroke-width='2.2' stroke-linecap='round'/>"
        "</svg><span>VidyaPath</span></div>",
        unsafe_allow_html=True,
    )
    st.caption("Your personal course companion")
    st.divider()
    if is_authenticated:
        st.caption(f"Signed in as {st.session_state.auth_user['email']}")
    else:
        st.caption("Guest session · sign in to save your learning")
    st.divider()
    st.subheader("Your course files")
    upload_widget_key = f"course_file_uploader_{st.session_state.get('course_file_uploader_version', 0)}"
    uploads = st.file_uploader("Add PDFs or PowerPoint slides", type=["pdf", "pptx"], accept_multiple_files=True, key=upload_widget_key)
    st.divider()
    st.subheader("Tutor settings")
    language = st.selectbox("Answer language", ["English", "Hindi"])
    st.caption("VidyaPath sends your question and matching excerpts to Google Gemini when you request an answer or quiz. Your original files are not sent to Gemini. Google may use free-tier prompts to improve its products.")
    st.caption("Tutor and quiz requests send your question and matching text excerpts to Google Gemini. If you opt in, Gemini analyzes your recorded voice question, the handwritten PDF, or video at your public YouTube link. It also turns your tutor's written answer into speech when you choose to listen. Google may use free-tier prompts to improve its products.")

for key, default in (("passages", []), ("loaded_files", set()), ("stored_file_paths", set()), ("file_paths_by_name", {}), ("removed_upload_paths", set()), ("files_needing_ocr", set()), ("video_sources", {}), ("video_notice", None), ("voice_result", None), ("voice_recording_version", 0), ("answers", []), ("tutor_history", []), ("tutor_history_error", None), ("study_notes", []), ("study_notes_error", None), ("quiz_history", []), ("used_quiz_questions", {}), ("quiz_counter", 0), ("current_quiz", None), ("course_maps", None)):
    if key not in st.session_state:
        st.session_state[key] = default

if st.session_state.get("course_content_parser_version") != 2:
    st.session_state.passages = []
    st.session_state.loaded_files = set(st.session_state.video_sources)
    st.session_state.stored_file_paths = set()
    st.session_state.file_paths_by_name = {}
    st.session_state.files_needing_ocr = set()
    st.session_state.course_maps = None
    for source_name, source in st.session_state.video_sources.items():
        for segment in source.get("segments") or []:
            location = f"{segment['timestamp']} · {segment['topic']}"
            st.session_state.passages.extend(split_passage(segment["notes"], source_name, location))
    if is_authenticated:
        st.session_state.active_user_id = None
    st.session_state.course_content_parser_version = 2

if is_authenticated and st.session_state.get("active_user_id") != user_id:
    st.session_state.active_user_id = user_id
    st.session_state.passages = []
    st.session_state.loaded_files = set()
    st.session_state.stored_file_paths = set()
    st.session_state.file_paths_by_name = {}
    st.session_state.removed_upload_paths = set()
    st.session_state.files_needing_ocr = set()
    st.session_state.video_sources = {}
    st.session_state.video_notice = None
    st.session_state.voice_result = None
    st.session_state.voice_recording_version += 1
    st.session_state.answers = []
    st.session_state.tutor_history = []
    st.session_state.tutor_history_error = None
    st.session_state.study_notes = []
    st.session_state.study_notes_error = None
    st.session_state.used_quiz_questions = {}
    st.session_state.current_quiz = None
    st.session_state.course_map = None
    st.session_state.course_maps = None
    st.session_state.quiz_counter = 0
    try:
        st.session_state.quiz_history = load_quiz_history(supabase)
    except Exception:
        st.session_state.quiz_history = []
        st.warning("Progress storage isn't ready yet. The app owner needs to run the included Supabase setup script.")
    try:
        st.session_state.tutor_history = load_tutor_history(supabase, user_id)
    except Exception:
        st.session_state.tutor_history_error = (
            "Saved tutor history isn't set up yet. Run supabase_tutor_history.sql in the Supabase SQL Editor."
        )
    try:
        st.session_state.study_notes = load_study_notes(supabase, user_id)
    except Exception:
        st.session_state.study_notes_error = (
            "Private study notes aren't set up yet. Run supabase_study_notes.sql in the Supabase SQL Editor."
        )
    try:
        for source in load_video_sources(supabase, user_id):
            source_name = f"YouTube · {source['video_id']}"
            st.session_state.video_sources[source_name] = source
            st.session_state.loaded_files.add(source_name)
            for segment in source.get("segments") or []:
                location = f"{segment['timestamp']} · {segment['topic']}"
                st.session_state.passages.extend(split_passage(segment["notes"], source_name, location))
    except Exception:
        st.info("Video lessons will be saved to your account after the app owner runs the updated Supabase setup script.")
    try:
        stored_items = supabase.storage.from_("course-files").list(user_id)
        for item in stored_items or []:
            stored_name = item.get("name", "")
            if not stored_name or stored_name in (".emptyFolderPlaceholder",):
                continue
            path = f"{user_id}/{stored_name}"
            raw_file = supabase.storage.from_("course-files").download(path)
            name_match = re.match(r"^[0-9a-f]{16}_(.*)$", stored_name, flags=re.IGNORECASE)
            display_name = name_match.group(1) if name_match else stored_name.split("_", 1)[-1]
            extracted = read_file(display_name, raw_file)
            st.session_state.passages.extend(extracted)
            if display_name.lower().endswith(".pdf") and not extracted:
                st.session_state.files_needing_ocr.add(display_name)
            st.session_state.loaded_files.add(display_name)
            st.session_state.stored_file_paths.add(path)
            st.session_state.file_paths_by_name.setdefault(display_name, set()).add(path)
    except Exception:
        st.warning("Your private file library isn't ready yet. The app owner needs to run the included Supabase setup script.")

active_upload_paths = set()
active_paths_by_name: dict[str, set[str]] = {}
for upload in uploads or []:
    clean_name = re.sub(r"[^A-Za-z0-9._-]", "_", upload.name)
    digest = hashlib.sha256(upload.getvalue()).hexdigest()[:16]
    upload_path = f"{user_id or 'guest'}/{digest}_{clean_name}"
    active_upload_paths.add(upload_path)
    active_paths_by_name.setdefault(upload.name, set()).add(upload_path)
st.session_state.removed_upload_paths.intersection_update(active_upload_paths)

if uploads:
    for upload in uploads:
        clean_name = re.sub(r"[^A-Za-z0-9._-]", "_", upload.name)
        file_bytes = upload.getvalue()
        digest = hashlib.sha256(file_bytes).hexdigest()[:16]
        storage_path = f"{user_id or 'guest'}/{digest}_{clean_name}"
        if storage_path in st.session_state.removed_upload_paths:
            continue
        if storage_path not in st.session_state.stored_file_paths:
            try:
                extracted = read_file(upload.name, file_bytes)
                if is_authenticated:
                    supabase.storage.from_("course-files").upload(
                        storage_path,
                        file_bytes,
                        file_options={
                            "content-type": "application/pdf" if upload.name.lower().endswith(".pdf") else "application/vnd.openxmlformats-officedocument.presentationml.presentation",
                            "upsert": "false",
                        },
                    )
                st.session_state.passages.extend(extracted)
                st.session_state.loaded_files.add(upload.name)
                st.session_state.stored_file_paths.add(storage_path)
                st.session_state.file_paths_by_name.setdefault(upload.name, set()).add(storage_path)
                st.session_state.course_maps = None
                if not is_authenticated:
                    st.caption("Guest mode: this file is temporary. Log in before uploading to save it to your account.")
                if not extracted:
                    if upload.name.lower().endswith(".pdf"):
                        st.session_state.files_needing_ocr.add(upload.name)
                        st.info(f"{upload.name} has no selectable text. You can choose to transcribe it from the course library below.")
                    else:
                        st.warning(f"No selectable text found in {upload.name}.")
            except Exception as error:
                st.error(f"Couldn't add {upload.name}: {error}")

video_notice = st.session_state.pop("video_notice", None)
if video_notice:
    notice_level, notice_text = video_notice
    getattr(st, notice_level)(notice_text)

header_space, account_column = st.columns([9, 2])
with account_column:
    if is_authenticated:
        profile_email = st.session_state.auth_user.get("email", "")
        profile_initial = escape(profile_email[:1].upper() or "?")
        profile_tooltip = escape(profile_email or "Signed-in profile")
        st.markdown(
            f"<div class='profile-avatar-row'><div class='profile-avatar' title='{profile_tooltip}' aria-label='{profile_tooltip}'>{profile_initial}</div></div>",
            unsafe_allow_html=True,
        )
        if st.button("Sign out", key="header_sign_out", type="primary", use_container_width=True):
            sign_out_current_user()
    elif st.button("Log in →", key="workspace_login_button", type="primary", use_container_width=True):
        st.session_state.show_login = True
        st.rerun()

st.markdown(
    "<div class='hero'><div class='eyebrow'>Learn with your own material</div><h1>Study smarter with VidyaPath.</h1>"
    "<p>Ask questions, get clear explanations, and follow every answer back to your course sources.</p></div>",
    unsafe_allow_html=True,
)

metric1, metric2, metric3 = st.columns(3)
with metric1:
    st.markdown(f"<div class='card metric'>COURSE FILES<strong>{len(st.session_state.loaded_files)}</strong></div>", unsafe_allow_html=True)
with metric2:
    st.markdown(f"<div class='card metric'>SEARCHABLE PASSAGES<strong>{len(st.session_state.passages)}</strong></div>", unsafe_allow_html=True)
with metric3:
    st.markdown("<div class='card metric'>YOUR STUDY MODE<strong>Tutor + practice</strong></div>", unsafe_allow_html=True)

st.markdown("<h2 style='text-align:center; margin-top:1.7rem;'>Your course library</h2>", unsafe_allow_html=True)
if st.session_state.loaded_files:
    library_files = sorted(st.session_state.loaded_files)
    library_columns = st.columns(min(len(library_files), 3))
    for index, filename in enumerate(library_files):
        with library_columns[index % len(library_columns)]:
            with st.container(border=True):
                st.write(f"📄 {filename}")
                if st.button("Remove", key=f"remove_file_{hashlib.sha1(filename.encode()).hexdigest()[:12]}"):
                    matching_paths = set(st.session_state.file_paths_by_name.get(filename, set()))
                    matching_paths.update(active_paths_by_name.get(filename, set()))
                    video_source = st.session_state.video_sources.get(filename)
                    try:
                        if is_authenticated:
                            if video_source:
                                delete_video_source(supabase, user_id, video_source["video_id"])
                            elif not matching_paths:
                                raise ValueError("This file is no longer in your saved library. Refresh the page and try again.")
                            else:
                                saved_paths = matching_paths.intersection(st.session_state.stored_file_paths)
                                if saved_paths:
                                    supabase.storage.from_("course-files").remove(list(saved_paths))
                        st.session_state.removed_upload_paths.update(matching_paths)
                        st.session_state.stored_file_paths.difference_update(matching_paths)
                        st.session_state.file_paths_by_name.pop(filename, None)
                        st.session_state.video_sources.pop(filename, None)
                        st.session_state.loaded_files.discard(filename)
                        st.session_state.passages = [p for p in st.session_state.passages if p.filename != filename]
                        st.session_state.answers = []
                        st.session_state.current_quiz = None
                        st.session_state.course_maps = None
                        st.rerun()
                    except Exception as error:
                        st.error(f"Couldn't remove the file: {error}")
                video_source = st.session_state.video_sources.get(filename)
                if video_source:
                    st.caption(video_source["title"])
                    st.link_button("Open lecture ↗", video_source["url"], key=f"open_video_{hashlib.sha1(filename.encode()).hexdigest()[:12]}")
                if filename in st.session_state.files_needing_ocr:
                    st.caption("This scanned PDF has no selectable text. Handwriting transcription uses Gemini's PDF reading and may make mistakes.")
                    consent_key = f"handwriting_consent_{hashlib.sha1(filename.encode()).hexdigest()[:12]}"
                    consent = st.checkbox(
                        "I agree to send this PDF to Google Gemini for transcription.",
                        key=consent_key,
                    )
                    if st.button(
                        "Read handwritten PDF",
                        key=f"transcribe_pdf_{hashlib.sha1(filename.encode()).hexdigest()[:12]}",
                        type="primary",
                        disabled=not consent,
                    ):
                        api_key = get_gemini_api_key()
                        if not api_key:
                            st.error("Handwriting transcription isn't connected yet. The app owner needs to add the Gemini key to the app's private configuration.")
                        else:
                            pdf_bytes = next(
                                (upload.getvalue() for upload in (uploads or []) if upload.name == filename),
                                None,
                            )
                            if pdf_bytes is None and is_authenticated:
                                candidate_paths = set(st.session_state.file_paths_by_name.get(filename, set()))
                                candidate_paths.intersection_update(st.session_state.stored_file_paths)
                                for saved_path in candidate_paths:
                                    try:
                                        pdf_bytes = supabase.storage.from_("course-files").download(saved_path)
                                        break
                                    except Exception:
                                        continue
                            if pdf_bytes is None:
                                st.error("This temporary guest PDF is no longer available. Upload it again, then choose transcription.")
                            else:
                                try:
                                    with st.spinner("Reading the handwritten pages…"):
                                        transcribed = transcribe_handwritten_pdf(pdf_bytes, api_key, filename)
                                    st.session_state.passages.extend(transcribed)
                                    st.session_state.files_needing_ocr.discard(filename)
                                    st.session_state.answers = []
                                    st.session_state.current_quiz = None
                                    st.session_state.course_maps = None
                                    st.session_state.pop(consent_key, None)
                                    st.rerun()
                                except Exception as error:
                                    st.error(f"Couldn't transcribe {filename}: {error}")
else:
    _, library_center, _ = st.columns([1, 2, 1])
    with library_center:
        st.info("Upload a PDF or PowerPoint from the left panel to start your course library.")

main, library = st.columns([2.7, 1], gap="large")
with main:
    api_key = get_gemini_api_key()
    # A keyed selector (not st.tabs) so the learner stays on the same section after the page refreshes.
    # st.tabs jumps back to the first tab whenever content above it changes, e.g. when a video is added to the library.
    section_names = ["Ask your tutor", "Practice quiz", "Your progress", "Study notes", "Course map", "Video lessons"]
    active_section = st.radio("Section", section_names, horizontal=True, key="active_section", label_visibility="collapsed")
    # Streamlit forgets inputs that are not drawn, so keep what the learner typed in the sections that are hidden right now.
    for _widget_key in list(st.session_state.keys()):
        if _widget_key in {"voice_processing_consent", "study_note_text", "youtube_video_url", "youtube_video_consent",
                           "quiz_topic", "tutor_question"} or re.fullmatch(r"quiz_\d+_\d+", str(_widget_key)):
            st.session_state[_widget_key] = st.session_state[_widget_key]
    if active_section == "Ask your tutor":
        st.markdown("Use a question about something in your uploaded course material.")
        if st.session_state.get("pending_tutor_voice_question"):
            st.session_state.tutor_question = st.session_state.pop("pending_tutor_voice_question")
        with st.expander("🎙️ Ask by voice"):
            voice_consent = st.checkbox(
                "I agree to send my recorded question and tutor answer text to Google Gemini.",
                key="voice_processing_consent",
            )
            voice_recording_key = f"voice_question_recording_{st.session_state.voice_recording_version}"
            voice_recording = st.audio_input("Record your question", sample_rate=16000, key=voice_recording_key)
            if voice_recording is not None:
                st.audio(voice_recording, format="audio/wav")
            if st.button(
                "Transcribe into my question",
                key="transcribe_tutor_voice_question",
                type="primary",
                use_container_width=True,
                disabled=not (voice_consent and voice_recording is not None),
            ):
                if not api_key:
                    st.warning("Voice transcription isn't connected yet. The app owner needs to add the Gemini key to the app's private configuration.")
                elif len(voice_recording.getvalue()) > 20 * 1024 * 1024:
                    st.warning("That recording is over 20 MB. Please record a shorter question.")
                elif not recording_contains_speech(voice_recording.getvalue()):
                    st.warning("I couldn't detect speech in that recording. Check the audio preview, allow microphone access, and try speaking closer to the microphone.")
                else:
                    try:
                        with st.spinner("Transcribing your recording…"):
                            st.session_state.pending_tutor_voice_question = transcribe_voice_question(
                                voice_recording.getvalue(), voice_recording.type, api_key
                            )
                        st.session_state.voice_recording_version += 1
                        st.rerun()
                    except Exception as error:
                        st.error(f"Couldn't transcribe that recording: {error}")
        question = st.text_input("Your question", placeholder="Ask a question about your uploaded course material", key="tutor_question")
        ask = st.button("✦  Explain with sources", type="primary", use_container_width=True)
        if ask:
            if not question.strip():
                st.warning("Type a question first.")
            elif not st.session_state.passages:
                st.warning("Upload a PDF or PowerPoint first.")
            elif not api_key:
                st.warning("The tutor is not connected yet. The app owner needs to add the Gemini key to the app's private configuration.")
            else:
                matches = retrieve(question, st.session_state.passages)
                if not matches:
                    st.info("I couldn't find a passage with matching terms. Try a course term or a person's full name.")
                else:
                    with st.spinner("Reading the matching course passages…"):
                        try:
                            answer = make_tutor_answer(question, matches, api_key, language)
                            st.session_state.answers.insert(0, {"question": question, "answer": answer, "sources": matches, "audio": None})
                            history_entry = make_tutor_history_entry(question, answer, matches)
                            remember_tutor_interaction(history_entry, supabase, user_id)
                        except Exception as error:
                            st.error(f"Tutor request failed: {error}")

        if st.session_state.answers:
            result = st.session_state.answers[0]
            st.markdown("### VidyaPath explains")
            with st.container(border=True):
                st.markdown(result["answer"])
            if result.get("audio"):
                st.audio(result["audio"], format="audio/wav")
                st.caption("Your spoken answer is ready. Press ▶ on the player to hear it.")
            elif st.button(
                "🔊 Listen to the answer",
                key=f"tutor_answer_audio_{hashlib.sha1(result['question'].encode()).hexdigest()[:10]}",
                type="primary",
                disabled=not st.session_state.get("voice_processing_consent", False),
            ):
                try:
                    with st.spinner("Preparing your spoken answer…"):
                        result["audio"] = synthesize_answer_audio(result["answer"], api_key, language)
                    st.rerun()
                except Exception as error:
                    st.warning(f"Couldn't create spoken audio, but the written answer is ready: {error}")
            st.markdown("**Sources from your course**")
            for index, source in enumerate(result["sources"], start=1):
                st.caption(f"[{index}] {source.filename} · {source.location}")
                video_source = st.session_state.video_sources.get(source.filename)
                time_match = re.match(r"(\d{2,}:\d{2}:\d{2})", source.location)
                if video_source and time_match:
                    st.link_button(
                        f"▶ Watch at {time_match.group(1)}",
                        video_timestamp_url(video_source, time_match.group(1)),
                        key=f"answer_video_source_{index}_{hashlib.sha1((source.filename + source.location).encode()).hexdigest()[:10]}",
                    )
                with st.expander(f"View source passage [{index}]"):
                    st.write(source.text)

        st.markdown("### Saved tutor history")
        if is_authenticated:
            st.caption("Your questions, answers, and cited course sources are saved to your account.")
        else:
            st.caption("History is kept for this guest session. Sign in to save it to your account.")
        if st.session_state.tutor_history_error:
            st.info(st.session_state.tutor_history_error)
        history = st.session_state.tutor_history
        if history:
            for index, entry in enumerate(history[:20]):
                created = str(entry.get("created_at") or "")[:16].replace("T", " ")
                question_label = str(entry.get("question") or "Previous question").strip()
                if len(question_label) > 90:
                    question_label = question_label[:87] + "…"
                heading = f"{created} · {question_label} · #{index + 1}" if created else f"{question_label} · #{index + 1}"
                with st.expander(heading):
                    st.markdown(entry.get("answer") or "No saved answer.")
                    sources = entry.get("sources") or []
                    if sources:
                        st.markdown("**Course sources**")
                        for source_index, source in enumerate(sources, start=1):
                            st.caption(
                                f"[{source_index}] {source.get('filename', 'Course material')} · "
                                f"{source.get('location', 'Source passage')}"
                            )
        else:
            st.caption("Your previous tutor questions will appear here.")

    if active_section == "Practice quiz":
        st.markdown("Create a short quiz from your own course material. Questions include an answer explanation and source citation.")
        topic = st.text_input("What topic do you want to practise?", placeholder="Enter a topic from your course material", key="quiz_topic")
        matching_history = [entry for entry in st.session_state.quiz_history if entry["topic"].casefold() == topic.strip().casefold()]
        if matching_history:
            recent_accuracy = sum(entry["score"] for entry in matching_history) / sum(entry["total"] for entry in matching_history)
            if recent_accuracy < 0.5:
                difficulty = "easier review"
            elif recent_accuracy >= 0.8:
                difficulty = "more challenging"
            else:
                difficulty = "standard practice"
            st.caption(f"Based on your earlier {topic} scores, this set will use **{difficulty}** questions.")
        else:
            difficulty = "introductory practice"
            st.caption("Your first set starts with introductory questions. The next set adjusts based on your score.")

        create_quiz = st.button("✦  Create my quiz", type="primary", use_container_width=True)
        if create_quiz:
            if not topic.strip():
                st.warning("Enter a topic first so the quiz can focus on the right course passages.")
            elif not st.session_state.passages:
                st.warning("Upload a PDF or PowerPoint first.")
            elif not api_key:
                st.warning("The tutor is not connected yet. The app owner needs to add the Gemini key to the app's private configuration.")
            else:
                matches = retrieve(topic, st.session_state.passages, limit=6)
                if not matches:
                    st.info("I couldn't find that topic in the uploaded material. Try a term that appears in the course files.")
                else:
                    with st.spinner("Building questions from your course sources…"):
                        try:
                            topic_key = topic.strip().casefold()
                            previous_questions = load_used_questions(supabase, user_id, topic_key) if is_authenticated else []
                            previous_questions.extend(st.session_state.used_quiz_questions.get(topic_key, []))
                            previous_questions = list(dict.fromkeys(previous_questions))
                            generated = generate_quiz(topic, matches, api_key, language, difficulty, previous_questions)
                            known_questions = {re.sub(r"[^a-z0-9]+", " ", text.lower()).strip() for text in previous_questions}
                            fresh_questions = []
                            for item in generated:
                                signature = re.sub(r"[^a-z0-9]+", " ", item["question"].lower()).strip()
                                if signature not in known_questions:
                                    fresh_questions.append(item)
                                    known_questions.add(signature)
                            generated = fresh_questions
                            if not generated:
                                st.warning("I couldn't find fresh questions for this topic. Try a more specific topic or another course section.")
                            else:
                                st.session_state.used_quiz_questions[topic_key] = previous_questions + [item["question"] for item in generated]
                                if is_authenticated:
                                    save_used_questions(supabase, user_id, topic_key, [item["question"] for item in generated])
                                st.session_state.current_quiz = {
                                    "topic": topic.strip(), "difficulty": difficulty, "questions": generated,
                                    "sources": matches, "submitted": False, "id": st.session_state.quiz_counter,
                                }
                                st.session_state.quiz_counter += 1
                        except Exception as error:
                            st.error(f"Couldn't create the quiz: {error}")

        quiz = st.session_state.current_quiz
        if quiz:
            st.markdown(f"### {quiz['topic']} · {quiz['difficulty'].title()}")
            if not quiz["submitted"]:
                for index, item in enumerate(quiz["questions"]):
                    st.markdown(f"**{index + 1}. {item['question']}**")
                    letters = list(item["options"])
                    st.radio(
                        f"Choose one answer for question {index + 1}", letters,
                        format_func=lambda letter, choices=item["options"]: f"{letter}. {choices[letter]}",
                        key=f"quiz_{quiz['id']}_{index}", label_visibility="collapsed", index=None,
                    )
                if st.button("Check my answers", type="primary", key=f"submit_quiz_{quiz['id']}"):
                    score = 0
                    responses = []
                    for index, item in enumerate(quiz["questions"]):
                        selected = st.session_state.get(f"quiz_{quiz['id']}_{index}")
                        is_correct = selected == item["correct"]
                        score += int(is_correct)
                        responses.append({"selected": selected, "correct": is_correct})
                    quiz["responses"] = responses
                    quiz["score"] = score
                    quiz["submitted"] = True
                    quiz_result = {"topic": quiz["topic"], "score": score, "total": len(quiz["questions"]), "difficulty": quiz["difficulty"]}
                    if is_authenticated:
                        save_quiz_result(supabase, user_id, quiz_result)
                    st.session_state.quiz_history.append(quiz_result)
                    st.rerun()
            else:
                st.success(f"You scored **{quiz['score']} out of {len(quiz['questions'])}**.")
                for index, (item, response) in enumerate(zip(quiz["questions"], quiz["responses"])):
                    label = "Correct" if response["correct"] else "Incorrect"
                    st.markdown(f"**{index + 1}. {label}: {item['question']}**")
                    if not response["correct"]:
                        st.write(f"Correct answer: **{item['correct']}. {item['options'][item['correct']]}**")
                    st.write(item["explanation"])
                    source = quiz["sources"][item["source"] - 1]
                    st.caption(f"Source: {source.filename} · {source.location}")
                    video_source = st.session_state.video_sources.get(source.filename)
                    time_match = re.match(r"(\d{2,}:\d{2}:\d{2})", source.location)
                    if video_source and time_match:
                        st.link_button(
                            f"▶ Watch at {time_match.group(1)}",
                            video_timestamp_url(video_source, time_match.group(1)),
                            key=f"quiz_video_source_{quiz['id']}_{index}",
                        )
                missed_questions = [
                    item.copy()
                    for item, response in zip(quiz["questions"], quiz["responses"])
                    if not response.get("correct")
                ]
                if missed_questions:
                    retry_column, new_quiz_column = st.columns(2)
                    if retry_column.button(
                        "Retry missed questions", type="primary", key=f"retry_missed_{quiz['id']}",
                        use_container_width=True,
                    ):
                        st.session_state.current_quiz = {
                            "topic": quiz["topic"],
                            "difficulty": "targeted review",
                            "questions": missed_questions,
                            "sources": quiz["sources"],
                            "submitted": False,
                            "id": st.session_state.quiz_counter,
                        }
                        st.session_state.quiz_counter += 1
                        st.rerun()
                    if new_quiz_column.button("Create another quiz", key=f"new_quiz_{quiz['id']}", use_container_width=True):
                        st.session_state.current_quiz = None
                        st.rerun()
                elif st.button("Create another quiz", key=f"new_quiz_{quiz['id']}"):
                    st.session_state.current_quiz = None
                    st.rerun()

    if active_section == "Your progress":
        if is_authenticated:
            st.markdown("Your quiz progress is saved to your signed-in account and available when you return.")
        else:
            st.markdown("You’re studying as a guest. Your progress lasts for this session; log in to save it to your account.")
        history = st.session_state.quiz_history
        if history:
            total_questions = sum(entry["total"] for entry in history)
            total_correct = sum(entry["score"] for entry in history)
            progress_a, progress_b, progress_c = st.columns(3)
            progress_a.metric("Quizzes completed", len(history))
            progress_b.metric("Questions answered", total_questions)
            progress_c.metric("Overall accuracy", f"{round(100 * total_correct / total_questions)}%" if total_questions else "—")
            topic_stats: dict[str, dict] = {}
            for entry in history:
                topic_key = entry["topic"].strip().casefold()
                stats = topic_stats.setdefault(topic_key, {
                    "name": entry["topic"].strip(), "correct": 0, "total": 0, "quizzes": 0,
                })
                stats["correct"] += entry["score"]
                stats["total"] += entry["total"]
                stats["quizzes"] += 1
            ranked_topics = sorted(
                topic_stats.values(),
                key=lambda item: item["correct"] / item["total"] if item["total"] else 0,
            )
            st.markdown("### Topic progress")
            st.caption("Accuracy below reflects your saved quiz history, not just your latest upload.")
            for item in ranked_topics:
                accuracy = item["correct"] / item["total"] if item["total"] else 0
                st.markdown(f"**{item['name']}** · {round(100 * accuracy)}% quiz accuracy")
                st.progress(accuracy)
        else:
            st.info("Complete your first quiz to see topic progress here.")

    if active_section == "Study notes":
        st.markdown("Save your own notes with the page, slide, or video timestamp they came from.")
        if st.session_state.study_notes_error:
            st.info(st.session_state.study_notes_error)

        source_options = sorted(
            {(passage.filename, passage.location) for passage in st.session_state.passages},
            key=lambda source: (source[0].casefold(), source[1].casefold()),
        )
        document_files = sorted(
            (filename for filename in st.session_state.loaded_files if filename.lower().endswith((".pdf", ".pptx"))),
            key=str.casefold,
        )
        if source_options or document_files:
            document_locations = {
                filename: sorted(
                    {
                        location for source_name, location in source_options
                        if source_name == filename
                        and re.search(r"\b(?:Page|Slide)\s+\d+\b", location, flags=re.IGNORECASE)
                    },
                    key=lambda location: int(re.search(r"\d+", location).group()) if re.search(r"\d+", location) else 0,
                )
                for filename in document_files
            }
            video_sources = [
                source for source in source_options
                if source[0] in st.session_state.video_sources
                and re.match(r"\d{2,}:\d{2}:\d{2}", source[1])
            ]
            pdf_col, video_col = st.columns(2)
            with pdf_col:
                with st.container(border=True):
                    st.markdown("### Open a course file")
                    if document_files:
                        selected_document = st.selectbox(
                            "Choose a PDF or PowerPoint",
                            document_files,
                            key="study_note_document_file",
                        )
                        is_pdf = selected_document.lower().endswith(".pdf")
                        matching_locations = document_locations.get(selected_document, [])
                        page_or_slide = re.search(r"\d+", matching_locations[0]).group() if matching_locations else "1"
                        if is_pdf and matching_locations:
                            selected_location = st.selectbox(
                                "Choose a PDF page",
                                matching_locations,
                                key="study_note_document_location",
                            )
                            page_or_slide = re.search(r"\d+", selected_location).group()

                        candidate_paths = set(st.session_state.file_paths_by_name.get(selected_document, set()))
                        candidate_paths.update(active_paths_by_name.get(selected_document, set()))
                        candidate_paths.intersection_update(st.session_state.stored_file_paths)
                        file_url = None
                        if is_authenticated and candidate_paths:
                            try:
                                saved_path = sorted(candidate_paths)[0]
                                signed_response = supabase.storage.from_("course-files").create_signed_url(saved_path, 3600)
                                signed_data = getattr(signed_response, "data", signed_response)
                                if isinstance(signed_data, dict):
                                    signed_url = signed_data.get("signedURL") or signed_data.get("signedUrl") or signed_data.get("signed_url")
                                    if signed_url:
                                        file_url = f"{signed_url}#page={page_or_slide}" if is_pdf else signed_url
                            except Exception as error:
                                st.info(f"Couldn't open this saved file right now: {error}")

                        if not is_pdf:
                            if file_url:
                                st.link_button("Open PowerPoint file ›", file_url, use_container_width=True)
                            else:
                                st.info("Sign in to open the PowerPoint saved in your course library.")
                        elif file_url:
                            st.link_button(
                                f"Open PDF at page {page_or_slide} ›",
                                file_url,
                                use_container_width=True,
                            )
                        else:
                            st.info("Sign in to open the PDF saved in your course library.")
                    else:
                        st.caption("Your uploaded PDFs and PowerPoint files will appear here.")

            with video_col:
                with st.container(border=True):
                    st.markdown("### Open a video timestamp")
                    if video_sources:
                        selected_video_source = st.selectbox(
                            "Choose a video topic",
                            video_sources,
                            format_func=lambda source: f"{st.session_state.video_sources[source[0]]['title']} · {source[1]}",
                            key="study_note_video_open_source",
                        )
                        selected_video = st.session_state.video_sources[selected_video_source[0]]
                        video_time_match = re.match(r"(\d{2,}:\d{2}:\d{2})", selected_video_source[1])
                        if video_time_match:
                            st.link_button(
                                "Open video at timestamp ›",
                                video_timestamp_url(selected_video, video_time_match.group(1)),
                                use_container_width=True,
                            )
                    else:
                        st.caption("Add a video lesson to open its topic timestamps here.")

            if source_options:
                selected_note_source = st.selectbox(
                    "Attach note to a course source",
                    source_options,
                    format_func=lambda source: f"{source[0]} · {source[1]}",
                    key="study_note_source",
                )

                with st.form("create_study_note_form", clear_on_submit=True):
                    note_text = st.text_area(
                        "Your note",
                        placeholder="Write a reminder or key idea…",
                        max_chars=3000,
                        key="study_note_text",
                    )
                    save_note = st.form_submit_button("Save note", type="primary", use_container_width=True)
                if save_note:
                    if not note_text.strip():
                        st.warning("Write a note before saving it.")
                    else:
                        note_entry = {
                            "id": uuid4().hex,
                            "source_filename": selected_note_source[0],
                            "source_location": selected_note_source[1],
                            "note": note_text.strip(),
                            "created_at": datetime.now(timezone.utc).isoformat(),
                        }
                        if is_authenticated:
                            try:
                                note_entry = save_study_note(supabase, user_id, note_entry)
                                st.session_state.study_notes_error = None
                            except Exception:
                                st.session_state.study_notes_error = (
                                    "This note is saved for the current session, but account storage isn't ready. "
                                    "Run supabase_study_notes.sql in the Supabase SQL Editor to keep it after you return."
                                )
                        st.session_state.study_notes.insert(0, note_entry)
        elif not st.session_state.study_notes:
            st.info("Upload course material or add a video lesson before creating a source-linked note.")

        if st.session_state.study_notes:
            st.markdown("### Your saved notes")
            for note in st.session_state.study_notes:
                note_id = note.get("id")
                with st.container(border=True):
                    st.markdown(
                        f"**{note.get('source_filename', 'Course material')}** · "
                        f"{note.get('source_location', 'Source')}",
                    )
                    st.write(note.get("note", ""))
                    if st.button("Remove note", key=f"remove_study_note_{note_id}"):
                        delete_succeeded = True
                        if is_authenticated and str(note_id).isdigit():
                            try:
                                delete_study_note(supabase, user_id, int(note_id))
                            except Exception:
                                st.error("Couldn't remove this saved note. Please try again.")
                                delete_succeeded = False
                        if delete_succeeded:
                            st.session_state.study_notes = [
                                saved_note for saved_note in st.session_state.study_notes
                                if saved_note.get("id") != note_id
                            ]
                            st.rerun()
        elif source_options:
            st.caption("Your saved notes will appear here.")

    if active_section == "Course map":
        if not st.session_state.passages:
            st.info("Upload a PDF or PowerPoint first to build a course map from your material.")
        elif not api_key:
            st.warning("The course map uses the app's Gemini connection. The app owner needs to add the Gemini key to the private configuration.")
        else:
            if st.button("✦  Build my course map", type="primary", use_container_width=True):
                with st.spinner("Comparing your files and organizing the topics…"):
                    try:
                        st.session_state.course_maps = generate_course_maps(st.session_state.passages, api_key)
                    except Exception as error:
                        st.error(f"Couldn't build the course map: {error}")
            course_maps = st.session_state.course_maps
            if isinstance(course_maps, dict):
                # Discard the previous single-map format during the upgrade.
                st.session_state.course_maps = None
                course_maps = None
            if course_maps:
                if len(course_maps) == 1:
                    st.graphviz_chart(course_map_graph(course_maps[0]), use_container_width=True)
                else:
                    separate_map_tabs = st.tabs([item["course"] for item in course_maps])
                    for map_tab_item, course_map_item in zip(separate_map_tabs, course_maps):
                        with map_tab_item:
                            st.graphviz_chart(course_map_graph(course_map_item), use_container_width=True)

    if active_section == "Video lessons":
        st.markdown("Add a public YouTube lecture to build a timestamped topic outline. The outline becomes a course source for tutor answers, quizzes, and course maps.")
        youtube_url = st.text_input(
            "Public YouTube lecture URL",
            placeholder="https://www.youtube.com/watch?v=…",
            key="youtube_video_url",
        )
        video_consent = st.checkbox(
            "I agree to let Google Gemini analyze the video from this public YouTube link.",
            key="youtube_video_consent",
        )
        if st.button("✦  Create timestamped outline", type="primary", use_container_width=True, disabled=not video_consent):
            video_id = get_youtube_video_id(youtube_url)
            if not video_id:
                st.warning("Enter a valid public YouTube video link. Private and unlisted videos aren't supported.")
            elif not api_key:
                st.warning("Video analysis isn't connected yet. The app owner needs to add the Gemini key to the app's private configuration.")
            else:
                try:
                    with st.spinner("Listening through the lecture and marking its topics…"):
                        source = analyze_youtube_video(video_id, api_key)
                    source_name = f"YouTube · {video_id}"
                    previous_source = st.session_state.video_sources.get(source_name)
                    if previous_source:
                        st.session_state.passages = [p for p in st.session_state.passages if p.filename != source_name]
                    st.session_state.video_sources[source_name] = source
                    st.session_state.loaded_files.add(source_name)
                    for segment in source["segments"]:
                        location = f"{segment['timestamp']} · {segment['topic']}"
                        st.session_state.passages.extend(split_passage(segment["notes"], source_name, location))
                    st.session_state.answers = []
                    st.session_state.current_quiz = None
                    st.session_state.course_maps = None
                    if is_authenticated:
                        try:
                            save_video_source(supabase, user_id, source)
                            notice = ("success", f"Created {len(source['segments'])} timestamped topics from {source['title']} and saved them to your account.")
                        except Exception:
                            notice = ("warning", "The outline is ready for this session, but saving it to your account needs the video setup SQL. Run supabase_video_sources.sql in Supabase first.")
                    else:
                        notice = ("info", f"Created {len(source['segments'])} timestamped topics from {source['title']}. Guest outlines are temporary; sign in before analyzing to save one to your account.")
                    st.session_state.video_notice = notice
                    st.rerun()
                except Exception as error:
                    st.error(f"Couldn't analyze this video: {error}")
        if st.session_state.video_sources:
            st.markdown("### Timestamped topics")
            video_names = list(st.session_state.video_sources)
            selected_video_name = st.selectbox(
                "Choose a video",
                video_names,
                format_func=lambda name: st.session_state.video_sources[name]["title"],
                key="timestamp_video_selector",
            )
            selected_video = st.session_state.video_sources[selected_video_name]
            for index, segment in enumerate(selected_video["segments"]):
                st.link_button(
                    f"▶  {segment['timestamp']} · {segment['topic']}",
                    video_timestamp_url(selected_video, segment["timestamp"]),
                    key=f"video_segment_{selected_video['video_id']}_{index}",
                    use_container_width=True,
                )
                st.caption(segment["notes"])


with library:
    st.subheader("Your learning path")
    st.markdown(
        "<div class='card'>"
        "<div class='path-step'><span class='path-icon' style='background:linear-gradient(135deg,#16a34a,#13a89e)'>✓</span><span><b>Find course sources</b><small>Upload PDFs and slides</small></span></div>"
        "<div class='path-step'><span class='path-icon' style='background:linear-gradient(135deg,#4361ee,#6541d9)'>✦</span><span><b>Ask the grounded tutor</b><small>Answers cite your material</small></span></div>"
        "<div class='path-step'><span class='path-icon' style='background:linear-gradient(135deg,#f59e0b,#ef4444)'>◷</span><span><b>Build quizzes and mastery</b><small>Practice and retry misses</small></span></div>"
        "<div class='path-step'><span class='path-icon' style='background:linear-gradient(135deg,#0ea5e9,#13a89e)'>⌘</span><span><b>Follow your course map</b><small>See how topics connect</small></span></div>"
        "</div>",
        unsafe_allow_html=True,
    )
    st.caption("Record a question or listen to its grounded answer directly in Ask your tutor.")