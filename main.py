import asyncio
import hashlib
import io
import json
import os
import re
import tempfile
import time
import urllib.parse
import urllib.request
import uuid
from typing import Optional

import docx
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt
from fastapi import FastAPI, File, Form, Header, HTTPException, Request, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, HTMLResponse, RedirectResponse, StreamingResponse
from google import genai
from google.genai import types
from pydantic import BaseModel
from supabase import Client, create_client

app = FastAPI(title="AI Quiz Generator API")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
MEDIA_DIR = os.path.join(BASE_DIR, "media_exports")
os.makedirs(MEDIA_DIR, exist_ok=True)

GEMINI_API_KEY = os.getenv("GEMINI_API_KEY", "").strip()
SUPABASE_URL = os.getenv("SUPABASE_URL", "https://gzeayqtnqwmxuwhccixl.supabase.co")
SUPABASE_KEY = os.getenv(
    "SUPABASE_KEY",
    "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Imd6ZWF5cXRucXdteHV3aGNjaXhsIiwicm9sZSI6InNlcnZpY2Vfcm9sZSIsImlhdCI6MTc5NDkxOTM3OCwiZXhwIjoyMTEwNDk1Mzc4fQ.TXWY-xSSj2YIeg_c42NA139kXw17y_fJUvpQo8Pklgc",
)

CLOUDINARY_CLOUD_NAME = os.getenv("CLOUDINARY_CLOUD_NAME", "").strip()
CLOUDINARY_API_KEY = os.getenv("CLOUDINARY_API_KEY", "").strip()
CLOUDINARY_API_SECRET = os.getenv("CLOUDINARY_API_SECRET", "").strip()
CLOUDINARY_FOLDER = "ai_quiz_media"

client = genai.Client(api_key=GEMINI_API_KEY)
supabase: Client = create_client(SUPABASE_URL, SUPABASE_KEY)

ALLOWED_GRADES = ["SD", "SMP", "SMA", "SMK", "Kuliah", "Umum"]


@app.get("/", response_class=HTMLResponse)
async def serve_home():
    index_path = os.path.join(BASE_DIR, "index.html")
    if os.path.exists(index_path):
        with open(index_path, "r", encoding="utf-8") as f:
            return f.read()
    return "<h3>File index.html tidak ditemukan di folder server.</h3>"


def _build_cloudinary_public_url(filename: str) -> Optional[str]:
    if not CLOUDINARY_CLOUD_NAME or not filename:
        return None
    safe_name = os.path.basename(filename.strip())
    stem, ext = os.path.splitext(safe_name)
    if not stem:
        return None
    cloud_enc = urllib.parse.quote(CLOUDINARY_CLOUD_NAME)
    public_id_enc = urllib.parse.quote(f"{CLOUDINARY_FOLDER}/{stem}", safe="/")
    ext_suffix = ext.lower() if ext else ""
    return f"https://res.cloudinary.com/{cloud_enc}/image/upload/{public_id_enc}{ext_suffix}"


def _upload_image_to_cloudinary(blob: bytes, public_id: str) -> Optional[str]:
    if not (CLOUDINARY_CLOUD_NAME and CLOUDINARY_API_KEY and CLOUDINARY_API_SECRET):
        return None

    timestamp = str(int(time.time()))
    params_to_sign = f"folder={CLOUDINARY_FOLDER}&public_id={public_id}&timestamp={timestamp}{CLOUDINARY_API_SECRET}"
    signature = hashlib.sha1(params_to_sign.encode("utf-8")).hexdigest()

    boundary = f"----WebKitFormBoundary{uuid.uuid4().hex}"
    crlf = "\r\n"

    def _part_field(name: str, value: str) -> bytes:
        return (
            f"--{boundary}{crlf}"
            f'Content-Disposition: form-data; name="{name}"{crlf}{crlf}'
            f"{value}{crlf}"
        ).encode("utf-8")

    body = bytearray()
    body.extend(_part_field("api_key", CLOUDINARY_API_KEY))
    body.extend(_part_field("timestamp", timestamp))
    body.extend(_part_field("folder", CLOUDINARY_FOLDER))
    body.extend(_part_field("public_id", public_id))
    body.extend(_part_field("signature", signature))
    body.extend(
        (
            f"--{boundary}{crlf}"
            f'Content-Disposition: form-data; name="file"; filename="{public_id}.png"{crlf}'
            f"Content-Type: application/octet-stream{crlf}{crlf}"
        ).encode("utf-8")
    )
    body.extend(blob)
    body.extend(f"{crlf}--{boundary}--{crlf}".encode("utf-8"))

    upload_url = f"https://api.cloudinary.com/v1_1/{urllib.parse.quote(CLOUDINARY_CLOUD_NAME)}/image/upload"
    req = urllib.request.Request(
        upload_url,
        data=bytes(body),
        headers={"Content-Type": f"multipart/form-data; boundary={boundary}"},
        method="POST",
    )

    try:
        with urllib.request.urlopen(req, timeout=30) as resp:
            payload = json.loads(resp.read().decode("utf-8"))
            return payload.get("secure_url")
    except Exception as cloud_err:
        print(f"Peringatan upload Cloudinary, fallback lokal: {cloud_err}")
        return None


def _normalize_media_url(raw_url: Optional[str], fallback_filename: Optional[str] = None) -> str:
    val = (raw_url or "").strip()
    fname = os.path.basename((fallback_filename or "").strip())

    if val.startswith("http://127.0.0.1:8000/media/") or val.startswith("http://localhost:8000/media/"):
        fname = os.path.basename(val)
        val = ""

    if val.startswith("http://") or val.startswith("https://"):
        return val

    if val.startswith("/media/"):
        fname = os.path.basename(val)

    if fname:
        local_path = os.path.join(MEDIA_DIR, fname)
        if os.path.exists(local_path):
            return f"/media/{urllib.parse.quote(fname)}"
        cloud_url = _build_cloudinary_public_url(fname)
        if cloud_url:
            return cloud_url
        return f"/media/{urllib.parse.quote(fname)}"

    return val


def _normalize_quiz_media_urls(quiz_obj: dict) -> dict:
    if not isinstance(quiz_obj, dict):
        return quiz_obj
    out = dict(quiz_obj)
    questions = out.get("questions")
    if not isinstance(questions, list):
        return out

    norm_questions = []
    for q in questions:
        if not isinstance(q, dict):
            norm_questions.append(q)
            continue
        item = dict(q)

        q_files = item.get("image_files") if isinstance(item.get("image_files"), list) else []
        q_urls = item.get("image_urls") if isinstance(item.get("image_urls"), list) else []
        max_q = max(len(q_files), len(q_urls))
        if max_q > 0:
            item["image_urls"] = [
                _normalize_media_url(
                    q_urls[i] if i < len(q_urls) else "",
                    q_files[i] if i < len(q_files) else "",
                )
                for i in range(max_q)
            ]

        if item.get("image_url") or item.get("image_file"):
            item["image_url"] = _normalize_media_url(item.get("image_url"), item.get("image_file"))

        opt_files = item.get("option_image_files") if isinstance(item.get("option_image_files"), list) else []
        opt_urls = item.get("option_image_urls") if isinstance(item.get("option_image_urls"), list) else []
        max_opt = max(len(opt_files), len(opt_urls))
        if max_opt > 0:
            item["option_image_urls"] = [
                _normalize_media_url(
                    opt_urls[i] if i < len(opt_urls) else "",
                    opt_files[i] if i < len(opt_files) else "",
                )
                for i in range(max_opt)
            ]

        if isinstance(item.get("pairs"), list):
            norm_pairs = []
            for p in item["pairs"]:
                if not isinstance(p, dict):
                    norm_pairs.append(p)
                    continue
                pd = dict(p)
                if pd.get("left_image_url") or pd.get("left_image_file"):
                    pd["left_image_url"] = _normalize_media_url(pd.get("left_image_url"), pd.get("left_image_file"))
                if pd.get("right_image_url") or pd.get("right_image_file"):
                    pd["right_image_url"] = _normalize_media_url(pd.get("right_image_url"), pd.get("right_image_file"))
                norm_pairs.append(pd)
            item["pairs"] = norm_pairs

        norm_questions.append(item)

    out["questions"] = norm_questions
    return out


def _normalize_grade(raw_grade: Optional[str]) -> str:
    if not raw_grade:
        return "Umum"
    text = raw_grade.strip().lower()
    if "smk" in text or "kejuruan" in text:
        return "SMK"
    if "sma" in text or "ma " in text or "madrasah aliyah" in text or any(k in text for k in ["kelas 10", "kelas 11", "kelas 12", "kelas x", "kelas xi", "kelas xii"]):
        return "SMA"
    if "smp" in text or "mts" in text or any(k in text for k in ["kelas 7", "kelas 8", "kelas 9", "kelas vii", "kelas viii", "kelas ix"]):
        return "SMP"
    if "sd" in text or "mi " in text or "madrasah ibtidaiyah" in text or any(k in text for k in ["kelas 1", "kelas 2", "kelas 3", "kelas 4", "kelas 5", "kelas 6"]):
        return "SD"
    if any(k in text for k in ["kuliah", "mahasiswa", "universitas", "kampus", "perguruan tinggi", "semester"]):
        return "Kuliah"
    for g in ALLOWED_GRADES:
        if g.lower() == text:
            return g
    return "Umum"


def _clean_filename(name: str) -> str:
    cleaned = re.sub(r"[^a-zA-Z0-9_-]+", "-", (name or "lembar-soal").strip()).strip("-")
    return cleaned or "lembar-soal"


def _clean_display_title(title: Optional[str], fallback: str = "Latihan Soal") -> str:
    t = (title or "").strip()
    if not t or "master" in t.lower():
        return fallback
    return t


def _strip_bracket_tags(text: str) -> str:
    cleaned = re.sub(r"\[[^\]]*\]", "", text or "")
    return re.sub(r"\s{2,}", " ", cleaned).strip()


def _normalize_passage_key(text: str) -> str:
    cleaned = _strip_bracket_tags(text or "").lower()
    return re.sub(r"[^a-z0-9]+", "", cleaned)


def _school_line_by_grade(grade: str) -> str:
    mapping = {
        "SD": "SEKOLAH DASAR NEGERI ......................",
        "SMP": "SEKOLAH MENENGAH PERTAMA NEGERI ......................",
        "SMA": "SEKOLAH MENENGAH ATAS NEGERI ......................",
        "SMK": "SEKOLAH MENENGAH KEJURUAN NEGERI ......................",
        "Kuliah": "FAKULTAS / PROGRAM STUDI ......................",
        "Umum": "LEMBAGA / INSTANSI ......................",
    }
    return mapping.get(grade, "SATUAN PENDIDIKAN ......................")


def _add_bottom_border(paragraph, sz: str = "12"):
    pPr = paragraph._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), sz)
    bottom.set(qn("w:space"), "4")
    bottom.set(qn("w:color"), "000000")
    pBdr.append(bottom)
    pPr.append(pBdr)


def _set_table_no_borders(table):
    tblPr = table._tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for side in ["top", "left", "bottom", "right", "insideH", "insideV"]:
        el = OxmlElement(f"w:{side}")
        el.set(qn("w:val"), "none")
        borders.append(el)
    tblPr.append(borders)


def _extract_images_from_xml_element(xml_element, doc_part, extracted_images: list) -> list:
    """Mengekstrak gambar inline/drawing dari elemen XML paragraf atau sel tabel."""
    found_tags = []
    blip_ns = "http://schemas.openxmlformats.org/drawingml/2006/main"
    embed_attr = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed"

    for blip in xml_element.findall(f".//{{{blip_ns}}}blip"):
        r_id = blip.get(embed_attr)
        if r_id and r_id in doc_part.related_parts:
            part = doc_part.related_parts[r_id]
            content_type = getattr(part, "content_type", "image/png") or "image/png"
            ext = ".png"
            if "jpeg" in content_type or "jpg" in content_type:
                ext = ".jpg"
            elif "gif" in content_type:
                ext = ".gif"
            elif "webp" in content_type:
                ext = ".webp"

            img_idx = len(extracted_images) + 1
            tag = f"[IMG_{img_idx}]"
            public_id = f"qimg_{uuid.uuid4().hex[:10]}_{img_idx}"
            filename = f"{public_id}{ext}"
            save_path = os.path.join(MEDIA_DIR, filename)

            with open(save_path, "wb") as img_f:
                img_f.write(part.blob)

            cloud_url = _upload_image_to_cloudinary(part.blob, public_id)
            final_url = cloud_url if cloud_url else f"/media/{filename}"

            extracted_images.append({
                "tag": tag,
                "filename": filename,
                "url": final_url,
                "content_type": content_type,
                "blob": part.blob,
            })
            found_tags.append(tag)
    return found_tags


def _extract_docx_content(doc: docx.Document, include_media: bool = False):
    """Membaca paragraf dan tabel secara berurutan sesuai posisi aslinya di dokumen."""
    blocks = []
    extracted_images = []

    for child in doc.element.body.iterchildren():
        local_name = child.tag.split("}")[-1]

        if local_name == "p":
            p = docx.text.paragraph.Paragraph(child, doc)
            text = (p.text or "").strip()
            img_tags = _extract_images_from_xml_element(child, doc.part, extracted_images) if include_media else []

            if text or img_tags:
                combined = text
                if img_tags:
                    combined = f"{text} {' '.join(img_tags)}".strip()
                blocks.append(combined)

        elif local_name == "tbl":
            tbl = docx.table.Table(child, doc)
            rows_data = []
            for row in tbl.rows:
                row_cells = []
                for cell in row.cells:
                    cell_text = " ".join([cp.text.strip() for cp in cell.paragraphs if cp.text.strip()])
                    c_imgs = _extract_images_from_xml_element(cell._tc, doc.part, extracted_images) if include_media else []
                    if c_imgs:
                        cell_text = f"{cell_text} {' '.join(c_imgs)}".strip()
                    row_cells.append(cell_text)
                if any(row_cells):
                    rows_data.append(row_cells)

            if rows_data and include_media:
                md_lines = ["| " + " | ".join(r) + " |" for r in rows_data]
                if len(rows_data) > 1:
                    sep = "| " + " | ".join(["---"] * len(rows_data[0])) + " |"
                    md_lines.insert(1, sep)
                blocks.append("[TABEL]\n" + "\n".join(md_lines) + "\n[/TABEL]")
            elif rows_data:
                blocks.extend([" | ".join([c for c in r if c]) for r in rows_data])

    return "\n".join(blocks), extracted_images


def _generate_single_json(prompt: str, extra_parts: list = None, temperature: float = 0.35, max_tokens: int = 32768) -> dict:
    contents = [prompt]
    if extra_parts:
        contents.extend(extra_parts)

    response = client.models.generate_content(
        model="gemini-3.1-flash-lite",
        contents=contents,
        config=types.GenerateContentConfig(
            response_mime_type="application/json",
            temperature=temperature,
            max_output_tokens=max_tokens,
        ),
    )
    raw_text = (response.text or "").strip()
    if raw_text.startswith("```"):
        raw_text = re.sub(r"^```(?:json)?\s*", "", raw_text)
        raw_text = re.sub(r"\s*```$", "", raw_text)
    return json.loads(raw_text)


@app.get("/media/{filename}")
async def serve_media_file(filename: str):
    safe_name = os.path.basename(filename)
    file_path = os.path.join(MEDIA_DIR, safe_name)
    if os.path.exists(file_path):
        return FileResponse(file_path)

    cloud_url = _build_cloudinary_public_url(safe_name)
    if cloud_url:
        return RedirectResponse(url=cloud_url, status_code=307)

    raise HTTPException(status_code=404, detail="File gambar tidak ditemukan.")


@app.post("/api/upload-quiz")
async def upload_and_process_quiz(
    file: UploadFile = File(...),
    is_premium: bool = Form(False),
    user_id: Optional[str] = Header(None),
    x_forwarded_for: Optional[str] = Header(None),
):
    if not file.filename or not file.filename.lower().endswith(".docx"):
        raise HTTPException(status_code=400, detail="Hanya file berformat .docx yang didukung.")

    effective_premium = is_premium
    if user_id:
        try:
            u_res = supabase.table("users").select("is_premium").eq("id", user_id).execute()
            if u_res.data:
                effective_premium = bool(u_res.data[0].get("is_premium", False))
        except Exception:
            pass

    variation_count = 2 if effective_premium else 1

    temp_path = None
    try:
        file_bytes = await file.read()

        with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as temp_file:
            temp_file.write(file_bytes)
            temp_path = temp_file.name

        doc = docx.Document(temp_path)
        extracted_text, extracted_images = _extract_docx_content(doc, include_media=effective_premium)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Gagal membaca file Word: {str(e)}")
    finally:
        if temp_path and os.path.exists(temp_path):
            try:
                os.remove(temp_path)
            except Exception:
                pass

    if not extracted_text.strip() and not extracted_images:
        raise HTTPException(status_code=400, detail="Dokumen Word kosong atau teks tidak terbaca.")

    media_instruction = """
    - Jika ada tabel pada soal, sertakan dalam array "table_rows" (contoh: [["Kolom 1", "Kolom 2"], ["Data A", "Data B"]]).
    - Jika ada tag gambar seperti [IMG_1], [IMG_2], dll:
      * Bila gambar berada di soal (jawaban teks), isi "image_tag": "[IMG_1]" dan sertakan semua tag gambar soal di "image_tags": ["[IMG_1]"] (bila ada 2 gambar atau lebih di soal yang sama, isi semuanya di "image_tags": ["[IMG_1]", "[IMG_2]"]).
      * Bila jawaban berupa gambar (opsi PG bergambar), isi "option_images": ["[IMG_1]", "[IMG_2]", "[IMG_3]", "[IMG_4]"] sesuai urutan opsi A-D.
      * Bila soal berbentuk menjodohkan (gambar vs teks atau gambar vs gambar), gunakan "type": "matching" dan isi "pairs": [{{"left": "Teks atau [IMG_1]", "right": "Teks atau [IMG_2]"}}].
    """ if effective_premium else """
    - Fokus pada soal berbasis teks (cerita, pilihan ganda, isian singkat, dan uraian). Jangan mengarang tag gambar baru.
    """

    master_prompt = f"""
    Anda adalah asisten pembuat soal pendidikan di Indonesia.
    Ekstrak SELURUH butir soal dari dokumen berikut ke dalam JSON murni. JANGAN meringkas atau memotong jumlah soal!

    ATURAN PENTING:
    1. Klasifikasikan "grade" ke SALAH SATU nilai baku berikut saja: "SD", "SMP", "SMA", "SMK", "Kuliah", atau "Umum".
    2. Susun "master_quiz" yang memuat SELURUH butir soal dari dokumen asli (bagian A Pilihan Ganda, bagian B Isian Singkat, Menjodohkan, maupun Uraian) dengan jumlah dan jenis yang SAMA PERSIS. Jangan berhenti di Pilihan Ganda saja!
    3. ATURAN KLASIFIKASI JENIS SOAL ("type"):
       - Gunakan "type": "multiple_choice" HANYA jika soal memiliki opsi pilihan jawaban (seperti A/B/C atau A/B/C/D).
       - Gunakan "type": "short_answer" untuk soal isian singkat, melengkapi titik-titik ("......" atau "______"), menyusun kata menjadi kalimat, atau mengubah kalimat aktif/pasif yang TIDAK memiliki opsi A/B/C/D. Set "options": [] dan isi kunci jawabannya pada "answer" serta "rubric".
       - Gunakan "type": "matching" untuk soal menjodohkan.
       - Gunakan "type": "essay" untuk soal uraian panjang berbasis teks/analisis. Set "options": [] dan isi pedoman jawaban pada "rubric" serta "answer".
    4. ATURAN SOAL CERITA / BACAAN MULTI-PARAGRAF:
       - Hanya isi "passage" pada soal yang memang secara langsung merujuk pada bacaan/cerita tersebut.
       - Jika satu cerita digunakan untuk beberapa soal (misalnya 3–5 soal), isi teks cerita yang SAMA PERSIS pada field "passage" untuk soal-soal tersebut agar sistem dapat mengelompokkannya otomatis.
       - Simpan SELURUH paragraf cerita/bacaan tersebut secara utuh (jangan diringkas).
       - Jika soal tidak memiliki bacaan/cerita khusus, isi "passage": "" (kosong).
    5. Jangan menyertakan label tanda kurung siku seperti "[PG]", "[Isian]", atau "[Uraian]" di dalam teks "question".
    {media_instruction}

    Struktur JSON WAJIB mengikuti format berikut:
    {{
        "metadata": {{
            "subject": "Mata pelajaran",
            "grade": "SD | SMP | SMA | SMK | Kuliah | Umum",
            "grade_detail": "Detail kelas jika ada (misal: Kelas 5 / Semester 1)",
            "topic": "Topik atau bab spesifik",
            "slug": "slug-ramah-seo-tanpa-spasi"
        }},
        "master_quiz": {{
            "title": "Latihan Soal Utama",
            "desc": "Paket soal acuan dari dokumen asli",
            "questions": [
                {{
                    "type": "multiple_choice",
                    "passage": "",
                    "question": "Teks pertanyaan bersih...",
                    "table_rows": [],
                    "image_tag": "",
                    "options": ["A. ...", "B. ...", "C. ...", "D. ..."],
                    "option_images": [],
                    "pairs": [],
                    "answer": "A",
                    "explanation": "Pembahasan singkat..."
                }}
            ]
        }}
    }}

    Teks Dokumen Soal:
    {extracted_text[:30000]}
    """

    image_parts = []
    if effective_premium and extracted_images:
        for img_item in extracted_images[:15]:
            image_parts.append(f"Referensi gambar {img_item['tag']}:")
            image_parts.append(
                types.Part.from_bytes(data=img_item["blob"], mime_type=img_item["content_type"])
            )

    try:
        master_bundle = await asyncio.to_thread(
            _generate_single_json,
            master_prompt,
            image_parts if image_parts else None,
            0.25,
            32768,
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Gagal memproses soal utama oleh AI: {str(e)}")

    meta = master_bundle.get("metadata") or {}
    normalized_grade = _normalize_grade(meta.get("grade"))
    meta["grade"] = normalized_grade

    raw_master_quiz = master_bundle.get("master_quiz") or {"title": "Latihan Soal Utama", "questions": []}
    master_questions_raw = raw_master_quiz.get("questions") or []

    type_counts = {}
    for mq in master_questions_raw:
        t_name = (mq.get("type") or "multiple_choice").strip()
        type_counts[t_name] = type_counts.get(t_name, 0) + 1
    type_summary_str = ", ".join([f"{k}: {v} soal" for k, v in type_counts.items()]) or f"{len(master_questions_raw)} soal"

    alternatives = []
    for var_idx in range(1, variation_count + 1):
        var_title = f"Paket Variasi {var_idx}"
        var_prompt = f"""
        Anda adalah pembuat variasi soal pendidikan di Indonesia.
        Berdasarkan daftar soal utama berikut, buatlah "{var_title}".

        ATURAN WAJIB:
        1. Jumlah soal pada "questions" HARUS SAMA PERSIS dengan soal utama ({len(master_questions_raw)} soal) dengan komposisi tipe WAJIB: {type_summary_str}. JANGAN mengubah soal short_answer/essay menjadi multiple_choice, dan JANGAN mengurangi jumlah soal!
        2. ACAK (shuffle) urutan nomor soal di dalam kelompok jenis soalnya masing-masing (kelompok multiple_choice diacak, short_answer diacak, matching diacak, essay diacak).
        3. UBAH redaksi/kalimat soal dan GANTI angka/nilai data hitungan agar tidak sama persis, namun tingkat kesulitannya tetap setara.
        4. Jika beberapa soal berbagi "passage" (cerita/bacaan) yang sama, pertahankan kelompok soal tersebut tetap berurutan dan gunakan teks "passage" hasil modifikasi yang SAMA PERSIS di antara soal-soal kelompok cerita itu (jangan dipindahkan ke soal lain yang tidak berkaitan).
        5. Pertahankan "type", "image_tag", "image_tags", "option_images", "pairs", dan "table_rows" yang sesuai dengan konteks soal/gambarnya. Untuk "short_answer" dan "essay", pastikan "options" tetap [] (kosong).
        6. Jangan menyertakan label kurung siku seperti "[PG]" di dalam teks "question".

        Format JSON:
        {{
            "title": "{var_title}",
            "desc": "Nomor diacak per tipe soal, redaksi & angka dimodifikasi, cerita utuh",
            "questions": []
        }}

        Daftar Soal Utama (JSON):
        {json.dumps(master_questions_raw, ensure_ascii=False)[:30000]}
        """
        try:
            var_bundle = await asyncio.to_thread(
                _generate_single_json,
                var_prompt,
                image_parts if image_parts else None,
                0.45 + (var_idx * 0.1),
                32768,
            )
            alt_questions = var_bundle.get("questions") or master_questions_raw
            alternatives.append({
                "title": var_title,
                "desc": var_bundle.get("desc") or "Nomor diacak per tipe soal, redaksi & angka dimodifikasi, cerita utuh",
                "questions": alt_questions,
            })
        except Exception as var_err:
            print(f"Peringatan {var_title} gagal detail, fallback dari soal utama: {var_err}")
            alternatives.append({
                "title": var_title,
                "desc": "Nomor diacak per tipe soal, redaksi & angka dimodifikasi, cerita utuh",
                "questions": master_questions_raw,
            })

    tag_to_meta = {item["tag"]: item for item in extracted_images}

    def _enrich_questions(q_list):
        enriched = []
        for q in q_list or []:
            item = dict(q)
            raw_tags = []
            if isinstance(item.get("image_tags"), list):
                raw_tags.extend([str(t).strip() for t in item.get("image_tags") if str(t).strip()])
            single_tag = (item.get("image_tag") or "").strip()
            if single_tag:
                raw_tags.append(single_tag)
            combined_scan_text = f"{item.get('question') or ''} {item.get('passage') or ''}"
            for stray_tag in re.findall(r"\[IMG_[^\]]+\]", combined_scan_text, flags=re.IGNORECASE):
                norm_tag = stray_tag.strip().upper()
                if norm_tag not in [t.upper() for t in raw_tags]:
                    raw_tags.append(norm_tag)

            ordered_q_tags = []
            seen_t = set()
            for t in raw_tags:
                t_up = t.upper()
                if t_up not in seen_t and t in tag_to_meta:
                    seen_t.add(t_up)
                    ordered_q_tags.append(t)

            item["image_tags"] = ordered_q_tags
            item["image_urls"] = [tag_to_meta[t]["url"] for t in ordered_q_tags]
            item["image_files"] = [tag_to_meta[t]["filename"] for t in ordered_q_tags]

            if ordered_q_tags:
                item["image_tag"] = ordered_q_tags[0]
                item["image_url"] = tag_to_meta[ordered_q_tags[0]]["url"]
                item["image_file"] = tag_to_meta[ordered_q_tags[0]]["filename"]

            raw_opt_imgs = item.get("option_images") or []
            resolved_opt_urls = []
            resolved_opt_files = []
            for t in raw_opt_imgs:
                t_str = (t or "").strip()
                if t_str in tag_to_meta:
                    resolved_opt_urls.append(tag_to_meta[t_str]["url"])
                    resolved_opt_files.append(tag_to_meta[t_str]["filename"])
                else:
                    resolved_opt_urls.append("")
                    resolved_opt_files.append("")
            item["option_image_urls"] = resolved_opt_urls
            item["option_image_files"] = resolved_opt_files

            raw_pairs = item.get("pairs") or []
            enriched_pairs = []
            for p in raw_pairs:
                p_dict = dict(p) if isinstance(p, dict) else {}
                l_val = (p_dict.get("left") or "").strip()
                r_val = (p_dict.get("right") or "").strip()
                if l_val in tag_to_meta:
                    p_dict["left_image_url"] = tag_to_meta[l_val]["url"]
                    p_dict["left_image_file"] = tag_to_meta[l_val]["filename"]
                if r_val in tag_to_meta:
                    p_dict["right_image_url"] = tag_to_meta[r_val]["url"]
                    p_dict["right_image_file"] = tag_to_meta[r_val]["filename"]
                enriched_pairs.append(p_dict)
            item["pairs"] = enriched_pairs
            enriched.append(item)
        return enriched

    master_quiz = {
        "title": _clean_display_title(raw_master_quiz.get("title"), fallback="Latihan Soal Utama"),
        "desc": raw_master_quiz.get("desc") or "Paket soal acuan dari dokumen asli",
        "questions": _enrich_questions(master_questions_raw),
    }

    for alt in alternatives:
        alt["questions"] = _enrich_questions(alt.get("questions"))

    ai_result = {
        "metadata": meta,
        "master_quiz": master_quiz,
        "alternatives": alternatives,
    }

    total_questions = len(master_quiz.get("questions") or [])

    archive_id = None
    try:
        archive_payload = {
            "user_id": user_id if user_id else None,
            "filename": file.filename,
            "subject": meta.get("subject") or "Umum",
            "grade": normalized_grade,
            "topic": meta.get("topic") or "Latihan Soal",
            "total_questions": total_questions,
            "master_quiz": master_quiz,
            "alternatives": alternatives,
        }
        arch_res = supabase.table("document_archives").insert(archive_payload).execute()
        if arch_res.data:
            archive_id = arch_res.data[0].get("id")
    except Exception as db_err:
        print(f"Peringatan simpan arsip: {db_err}")

    return {
        "status": "success",
        "archive_id": archive_id,
        "tier": "premium" if effective_premium else "free",
        "data": ai_result,
    }


@app.get("/api/archives")
async def list_archives(grade: Optional[str] = None):
    try:
        query = (
            supabase.table("document_archives")
            .select("id, filename, subject, grade, topic, total_questions, created_at")
            .order("created_at", desc=True)
            .limit(200)
        )
        if grade and grade in ALLOWED_GRADES:
            query = query.eq("grade", grade)

        res = query.execute()
        rows = res.data or []

        counts = {g: 0 for g in ALLOWED_GRADES}
        for r in rows:
            g_norm = _normalize_grade(r.get("grade"))
            r["grade"] = g_norm
            counts[g_norm] = counts.get(g_norm, 0) + 1

        return {"status": "success", "counts": counts, "archives": rows}
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Gagal memuat arsip dokumen: {str(e)}")


@app.get("/api/archives/{archive_id}")
async def get_archive_detail(archive_id: str):
    try:
        res = supabase.table("document_archives").select("*").eq("id", archive_id).execute()
        if not res.data:
            raise HTTPException(status_code=404, detail="Arsip dokumen tidak ditemukan.")
        row = res.data[0]
        raw_master = row.get("master_quiz", {"title": "Latihan Soal Utama", "questions": []})
        raw_alts = row.get("alternatives", [])

        norm_master = _normalize_quiz_media_urls(raw_master)
        norm_alts = [_normalize_quiz_media_urls(a) for a in raw_alts] if isinstance(raw_alts, list) else []

        return {
            "status": "success",
            "archive": row,
            "data": {
                "metadata": {
                    "subject": row.get("subject", "Umum"),
                    "grade": _normalize_grade(row.get("grade")),
                    "topic": row.get("topic", "Latihan Soal"),
                    "slug": _clean_filename(row.get("filename", "latihan-soal")),
                },
                "master_quiz": norm_master,
                "alternatives": norm_alts,
            },
        }
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Gagal memuat detail arsip: {str(e)}")


class ExportDocxRequest(BaseModel):
    title: str = "Lembar Latihan Soal"
    subject: str = "Umum"
    grade: str = "Umum"
    grade_detail: Optional[str] = ""
    topic: Optional[str] = "Latihan Soal"
    filename_slug: Optional[str] = "lembar-soal"
    quiz_data: dict


def _insert_image_if_exists(doc: docx.Document, image_file: Optional[str], image_url: Optional[str] = None, width_cm: float = 6.5):
    img_source = None
    if image_file:
        safe_name = os.path.basename(image_file)
        img_path = os.path.join(MEDIA_DIR, safe_name)
        if os.path.exists(img_path):
            img_source = img_path

    if img_source is None:
        candidate_url = _normalize_media_url(image_url, image_file)
        if candidate_url.startswith("http://") or candidate_url.startswith("https://"):
            try:
                with urllib.request.urlopen(candidate_url, timeout=15) as resp:
                    img_source = io.BytesIO(resp.read())
            except Exception:
                img_source = None

    if img_source is not None:
        try:
            p_img = doc.add_paragraph()
            p_img.paragraph_format.left_indent = Cm(0.75)
            p_img.paragraph_format.space_before = Pt(2)
            p_img.paragraph_format.space_after = Pt(4)
            p_img.add_run().add_picture(img_source, width=Cm(width_cm))
        except Exception:
            pass


def _insert_table_if_exists(doc: docx.Document, table_rows):
    if not isinstance(table_rows, list) or len(table_rows) == 0:
        return
    col_count = max(len(r) for r in table_rows if isinstance(r, list))
    if col_count <= 0:
        return

    tbl = doc.add_table(rows=len(table_rows), cols=col_count)
    tbl.style = "Table Grid"
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER

    for r_idx, row_vals in enumerate(table_rows):
        row = tbl.rows[r_idx]
        for c_idx in range(col_count):
            val = str(row_vals[c_idx]) if isinstance(row_vals, list) and c_idx < len(row_vals) else ""
            cell_p = row.cells[c_idx].paragraphs[0]
            cell_p.paragraph_format.space_before = Pt(1)
            cell_p.paragraph_format.space_after = Pt(1)
            run = cell_p.add_run(_strip_bracket_tags(val))
            run.font.size = Pt(10)
            if r_idx == 0:
                run.bold = True

    p_sp = doc.add_paragraph()
    p_sp.paragraph_format.space_before = Pt(0)
    p_sp.paragraph_format.space_after = Pt(4)


@app.post("/api/export-docx")
async def export_quiz_to_docx(payload: ExportDocxRequest):
    try:
        doc = docx.Document()

        for section in doc.sections:
            section.page_width = Cm(21)
            section.page_height = Cm(33)
            section.top_margin = Cm(2.2)
            section.bottom_margin = Cm(2.2)
            section.left_margin = Cm(2.2)
            section.right_margin = Cm(2.2)

        style = doc.styles["Normal"]
        style.font.name = "Times New Roman"
        style.font.size = Pt(11)

        norm_grade = _normalize_grade(payload.grade)
        kelas_tampil = (payload.grade_detail or "").strip() or (norm_grade if norm_grade != "Umum" else ".........................")

        p_pemda = doc.add_paragraph("PEMERINTAH KABUPATEN / KOTA ........................................")
        p_pemda.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_pemda.paragraph_format.space_before = Pt(0)
        p_pemda.paragraph_format.space_after = Pt(1)
        p_pemda.runs[0].bold = True
        p_pemda.runs[0].font.size = Pt(11)

        p_dinas = doc.add_paragraph("DINAS PENDIDIKAN")
        p_dinas.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_dinas.paragraph_format.space_before = Pt(0)
        p_dinas.paragraph_format.space_after = Pt(1)
        p_dinas.runs[0].bold = True
        p_dinas.runs[0].font.size = Pt(12)

        p_sekolah = doc.add_paragraph(_school_line_by_grade(norm_grade))
        p_sekolah.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_sekolah.paragraph_format.space_before = Pt(0)
        p_sekolah.paragraph_format.space_after = Pt(1)
        p_sekolah.runs[0].bold = True
        p_sekolah.runs[0].font.size = Pt(12)

        p_alamat = doc.add_paragraph("Alamat: ............................................................................................")
        p_alamat.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_alamat.paragraph_format.space_before = Pt(0)
        p_alamat.paragraph_format.space_after = Pt(6)
        p_alamat.runs[0].font.size = Pt(10)
        _add_bottom_border(p_alamat, sz="18")

        p_asesmen = doc.add_paragraph("ASESMEN SUMATIF AKHIR SEMESTER (ASAS)")
        p_asesmen.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_asesmen.paragraph_format.space_before = Pt(4)
        p_asesmen.paragraph_format.space_after = Pt(1)
        p_asesmen.runs[0].bold = True
        p_asesmen.runs[0].font.size = Pt(12)

        p_tp = doc.add_paragraph("TAHUN PELAJARAN 2026/2027")
        p_tp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_tp.paragraph_format.space_before = Pt(0)
        p_tp.paragraph_format.space_after = Pt(8)
        p_tp.runs[0].bold = True
        p_tp.runs[0].font.size = Pt(11)

        id_table = doc.add_table(rows=3, cols=2)
        id_table.alignment = WD_TABLE_ALIGNMENT.CENTER
        _set_table_no_borders(id_table)

        left_items = [
            f"Mata Pelajaran : {payload.subject or '.........................'}",
            "Hari / Tanggal  : .........................",
            f"Kelas / Semester: {kelas_tampil}",
        ]
        right_items = [
            "Nama Siswa  : .........................",
            "No. Absen   : .........................",
            "Nilai       : .........................",
        ]

        for r_idx in range(3):
            row = id_table.rows[r_idx]
            row.cells[0].width = Cm(8.5)
            row.cells[1].width = Cm(8.1)

            c_left = row.cells[0].paragraphs[0]
            c_left.paragraph_format.space_before = Pt(1)
            c_left.paragraph_format.space_after = Pt(1)
            run_l = c_left.add_run(left_items[r_idx])
            run_l.font.size = Pt(10.5)

            c_right = row.cells[1].paragraphs[0]
            c_right.paragraph_format.space_before = Pt(1)
            c_right.paragraph_format.space_after = Pt(1)
            run_r = c_right.add_run(right_items[r_idx])
            run_r.font.size = Pt(10.5)

        p_divider = doc.add_paragraph()
        p_divider.paragraph_format.space_before = Pt(2)
        p_divider.paragraph_format.space_after = Pt(10)
        _add_bottom_border(p_divider, sz="12")

        questions = payload.quiz_data.get("questions") or []
        pg_list, short_list, matching_list, essay_list = [], [], [], []

        for q in questions:
            q_type = (q.get("type") or "").lower().strip()
            opts = q.get("options") if isinstance(q.get("options"), list) else []
            if q_type == "multiple_choice" and len(opts) > 0:
                pg_list.append(q)
            elif q_type in ("short_answer", "isian", "fill_in_the_blank") or (q_type == "multiple_choice" and len(opts) == 0):
                short_list.append(q)
            elif q_type == "matching":
                matching_list.append(q)
            else:
                essay_list.append(q)

        def _cluster_by_passage(items: list) -> list:
            grouped_order = []
            visited = [False] * len(items)
            for i, itm in enumerate(items):
                if visited[i]:
                    continue
                visited[i] = True
                grouped_order.append(itm)
                p_key = _normalize_passage_key(itm.get("passage") or "")
                if p_key:
                    for j in range(i + 1, len(items)):
                        if not visited[j]:
                            other_key = _normalize_passage_key(items[j].get("passage") or "")
                            if other_key and (other_key == p_key or (len(p_key) > 40 and (p_key[:60] == other_key[:60]))):
                                visited[j] = True
                                grouped_order.append(items[j])
            return grouped_order

        pg_list = _cluster_by_passage(pg_list)
        short_list = _cluster_by_passage(short_list)
        matching_list = _cluster_by_passage(matching_list)
        essay_list = _cluster_by_passage(essay_list)

        ordered_questions = (
            [("pg", item) for item in pg_list]
            + [("short", item) for item in short_list]
            + [("matching", item) for item in matching_list]
            + [("essay", item) for item in essay_list]
        )

        printed_pg_header = False
        printed_short_header = False
        printed_matching_header = False
        printed_essay_header = False
        printed_passage_keys = set()
        answer_lines = []

        for idx, (group_type, q) in enumerate(ordered_questions, start=1):
            if group_type == "pg" and not printed_pg_header:
                p_inst = doc.add_paragraph("BERILAH TANDA SILANG (X) PADA HURUF A, B, C ATAU D PADA JAWABAN YANG PALING BENAR!")
                p_inst.paragraph_format.keep_with_next = True
                p_inst.paragraph_format.space_before = Pt(4)
                p_inst.paragraph_format.space_after = Pt(6)
                p_inst.runs[0].bold = True
                p_inst.runs[0].font.size = Pt(10.5)
                printed_pg_header = True

            elif group_type == "short" and not printed_short_header:
                p_inst = doc.add_paragraph("ISILAH PERTANYAAN-PERTANYAAN BERIKUT DENGAN JAWABAN SINGKAT DAN JELAS!")
                p_inst.paragraph_format.keep_with_next = True
                p_inst.paragraph_format.space_before = Pt(8)
                p_inst.paragraph_format.space_after = Pt(6)
                p_inst.runs[0].bold = True
                p_inst.runs[0].font.size = Pt(10.5)
                printed_short_header = True

            elif group_type == "matching" and not printed_matching_header:
                p_inst = doc.add_paragraph("JODOHKAN PERNYATAAN / GAMBAR DI BAGIAN KIRI DENGAN PASANGAN YANG SESUAI DI BAGIAN KANAN!")
                p_inst.paragraph_format.keep_with_next = True
                p_inst.paragraph_format.space_before = Pt(8)
                p_inst.paragraph_format.space_after = Pt(6)
                p_inst.runs[0].bold = True
                p_inst.runs[0].font.size = Pt(10.5)
                printed_matching_header = True

            elif group_type == "essay" and not printed_essay_header:
                p_inst = doc.add_paragraph("JAWABLAH PERTANYAAN DI BAWAH INI DENGAN LENGKAP!")
                p_inst.paragraph_format.keep_with_next = True
                p_inst.paragraph_format.space_before = Pt(8)
                p_inst.paragraph_format.space_after = Pt(6)
                p_inst.runs[0].bold = True
                p_inst.runs[0].font.size = Pt(10.5)
                printed_essay_header = True

            passage = (q.get("passage") or "").strip()
            p_key = _normalize_passage_key(passage)
            p_prefix = p_key[:80] if len(p_key) > 80 else p_key

            if passage and p_prefix and p_prefix not in printed_passage_keys:
                span_count = 1
                for next_idx in range(idx, len(ordered_questions)):
                    _, next_q = ordered_questions[next_idx]
                    next_key = _normalize_passage_key(next_q.get("passage") or "")
                    next_prefix = next_key[:80] if len(next_key) > 80 else next_key
                    if next_prefix == p_prefix:
                        span_count += 1
                    else:
                        break

                if span_count > 1:
                    p_ref = doc.add_paragraph(f"Bacalah teks berikut untuk menjawab soal nomor {idx} s.d. {idx + span_count - 1}!")
                    p_ref.paragraph_format.space_before = Pt(4)
                    p_ref.paragraph_format.space_after = Pt(2)
                    p_ref.paragraph_format.keep_with_next = True
                    p_ref.runs[0].italic = True
                    p_ref.runs[0].bold = True

                lines = [ln.strip() for ln in passage.split("\n") if ln.strip()]
                for l_i, para_line in enumerate(lines):
                    clean_para = _strip_bracket_tags(para_line)
                    if clean_para:
                        p_passage = doc.add_paragraph()
                        p_passage.paragraph_format.space_before = Pt(2)
                        p_passage.paragraph_format.space_after = Pt(3)
                        p_passage.paragraph_format.keep_with_next = True
                        run_p = p_passage.add_run(clean_para)
                        if l_i == 0 and len(clean_para) <= 65 and not clean_para.endswith("."):
                            p_passage.alignment = WD_ALIGN_PARAGRAPH.CENTER
                            run_p.bold = True
                        else:
                            p_passage.paragraph_format.first_line_indent = Cm(0.75)

                printed_passage_keys.add(p_prefix)

            q_text = _strip_bracket_tags(q.get("question") or "")
            p_q = doc.add_paragraph()
            p_q.paragraph_format.keep_with_next = True
            p_q.paragraph_format.space_before = Pt(2)
            p_q.paragraph_format.space_after = Pt(2)
            p_q.add_run(f"{idx}. {q_text}")

            q_files_to_print = (
                q.get("image_files")
                if isinstance(q.get("image_files"), list) and q.get("image_files")
                else ([q.get("image_file")] if q.get("image_file") else [])
            )
            q_urls_to_print = (
                q.get("image_urls")
                if isinstance(q.get("image_urls"), list) and q.get("image_urls")
                else ([q.get("image_url")] if q.get("image_url") else [])
            )
            max_q_imgs = max(len(q_files_to_print), len(q_urls_to_print))
            for img_i in range(max_q_imgs):
                qf = q_files_to_print[img_i] if img_i < len(q_files_to_print) else None
                qu = q_urls_to_print[img_i] if img_i < len(q_urls_to_print) else None
                _insert_image_if_exists(doc, qf, qu, width_cm=7.0)

            _insert_table_if_exists(doc, q.get("table_rows"))

            if group_type == "pg":
                opts = q.get("options") if isinstance(q.get("options"), list) else []
                opt_files = q.get("option_image_files") if isinstance(q.get("option_image_files"), list) else []
                opt_urls = q.get("option_image_urls") if isinstance(q.get("option_image_urls"), list) else []
                for o_idx, opt in enumerate(opts):
                    clean_opt = _strip_bracket_tags(str(opt))
                    p_opt = doc.add_paragraph(clean_opt)
                    p_opt.paragraph_format.left_indent = Cm(0.75)
                    p_opt.paragraph_format.space_before = Pt(1)
                    p_opt.paragraph_format.space_after = Pt(1)
                    of = opt_files[o_idx] if o_idx < len(opt_files) else None
                    ou = opt_urls[o_idx] if o_idx < len(opt_urls) else None
                    if of or ou:
                        _insert_image_if_exists(doc, of, ou, width_cm=4.5)

                ans = q.get("answer") or "-"
                exp = _strip_bracket_tags(q.get("explanation") or "")
                answer_lines.append(f"{idx}. Jawaban: {ans}" + (f" — Pembahasan: {exp}" if exp else ""))

            elif group_type == "matching":
                pairs = q.get("pairs") if isinstance(q.get("pairs"), list) else []
                if pairs:
                    m_tbl = doc.add_table(rows=len(pairs) + 1, cols=2)
                    m_tbl.style = "Table Grid"
                    m_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
                    m_tbl.rows[0].cells[0].paragraphs[0].add_run("Bagian A").bold = True
                    m_tbl.rows[0].cells[1].paragraphs[0].add_run("Bagian B").bold = True
                    for p_i, pair_item in enumerate(pairs, start=1):
                        l_txt = _strip_bracket_tags(str(pair_item.get("left") or ""))
                        r_txt = _strip_bracket_tags(str(pair_item.get("right") or ""))
                        m_tbl.rows[p_i].cells[0].paragraphs[0].add_run(f"{p_i}. {l_txt}")
                        m_tbl.rows[p_i].cells[1].paragraphs[0].add_run(f"{chr(64 + p_i)}. {r_txt}")
                answer_lines.append(f"{idx}. Soal Menjodohkan — Pasangkan nomor sesuai kunci pasangan.")

            elif group_type == "short":
                p_ans_line = doc.add_paragraph("Jawab: ............................................................................................")
                p_ans_line.paragraph_format.left_indent = Cm(0.75)
                p_ans_line.paragraph_format.space_before = Pt(2)
                p_ans_line.paragraph_format.space_after = Pt(4)

                rubric = _strip_bracket_tags(q.get("rubric") or q.get("answer") or "-")
                answer_lines.append(f"{idx}. Kunci Jawaban: {rubric}")

            else:
                for _ in range(3):
                    p_line = doc.add_paragraph("....................................................................................................")
                    p_line.paragraph_format.left_indent = Cm(0.75)
                    p_line.paragraph_format.space_before = Pt(2)
                    p_line.paragraph_format.space_after = Pt(2)

                rubric = _strip_bracket_tags(q.get("rubric") or q.get("answer") or "-")
                answer_lines.append(f"{idx}. Rubrik / Kunci: {rubric}")

        doc.add_page_break()
        key_heading = doc.add_heading("KUNCI JAWABAN & PEDOMAN PENILAIAN", level=2)
        key_heading.runs[0].font.size = Pt(13)

        for line in answer_lines:
            p_ans = doc.add_paragraph(line)
            p_ans.paragraph_format.space_after = Pt(3)

        buffer = io.BytesIO()
        doc.save(buffer)
        buffer.seek(0)

        safe_name = f"{_clean_filename(payload.filename_slug or payload.title)}.docx"
        return StreamingResponse(
            buffer,
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            headers={"Content-Disposition": f'attachment; filename="{safe_name}"'},
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Gagal membuat file Word (.docx): {str(e)}")


class PublishQuizRequest(BaseModel):
    title: str = "Latihan Soal"
    subject: str = "Umum"
    grade: str = "Umum"
    topic: Optional[str] = "Latihan Soal"
    slug: str
    quiz_data: dict
    user_id: Optional[str] = None


@app.post("/api/publish-quiz")
async def publish_quiz(payload: PublishQuizRequest, request: Request):
    try:
        clean_title = _clean_display_title(payload.title, fallback=payload.topic or "Latihan Soal")
        normalized_quiz_data = _normalize_quiz_media_urls(payload.quiz_data)
        insert_data = {
            "title": clean_title,
            "subject": payload.subject,
            "grade": payload.grade,
            "topic": payload.topic,
            "slug": payload.slug,
            "quiz_data": normalized_quiz_data,
            "user_id": payload.user_id,
            "is_published": True,
        }

        response = supabase.table("quizzes").insert(insert_data).execute()

        if not response.data:
            raise HTTPException(status_code=500, detail="Gagal menyimpan kuis ke database.")

        saved_quiz = response.data[0]
        base_url = str(request.base_url).rstrip("/")

        return {
            "status": "success",
            "message": "Kuis berhasil dipublikasikan!",
            "quiz_id": saved_quiz["id"],
            "slug": saved_quiz["slug"],
            "share_url": f"{base_url}/quiz/{saved_quiz['slug']}",
        }

    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Terjadi kesalahan database: {str(e)}")


@app.get("/api/get-quiz/{slug}")
async def get_quiz_by_slug(slug: str):
    try:
        response = supabase.table("quizzes").select("*").eq("slug", slug).execute()
        if not response.data:
            raise HTTPException(status_code=404, detail="Kuis tidak ditemukan")

        row = response.data[0]
        if isinstance(row.get("quiz_data"), dict):
            row["quiz_data"] = _normalize_quiz_media_urls(row["quiz_data"])
        return row
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@app.get("/quiz/{slug}", response_class=HTMLResponse)
async def serve_quiz_page(slug: str):
    try:
        with open(os.path.join(BASE_DIR, "quiz_template.html"), "r", encoding="utf-8") as f:
            return f.read()
    except Exception:
        return "<h3>Template kuis tidak ditemukan di server.</h3>", 404