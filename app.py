from flask import Flask, render_template, request, jsonify
from PIL import Image, ExifTags
import PyPDF2
import docx
import os
from datetime import datetime
import openpyxl
from pymediainfo import MediaInfo

app = Flask(__name__)

# ----------------- Metadata Extraction ----------------- #
def convert_gps(coord, ref):
    d = coord[0][0] / coord[0][1]
    m = coord[1][0] / coord[1][1]
    s = coord[2][0] / coord[2][1]
    decimal = d + (m / 60.0) + (s / 3600.0)
    if ref in ['S', 'W']:
        decimal = -decimal
    return decimal

def gps_to_link(lat, lon):
    return f"https://www.google.com/maps?q={lat},{lon}"

def extract_image_metadata(path):
    metadata = {}
    try:
        image = Image.open(path)
        exif_data = image._getexif()
        if exif_data:
            gps_info = None
            for tag_id, value in exif_data.items():
                tag = ExifTags.TAGS.get(tag_id, tag_id)
                if tag == "GPSInfo":
                    gps_info = value
                else:
                    metadata[tag] = str(value)
            if gps_info:
                gps_tags = {}
                for key in gps_info.keys():
                    name = ExifTags.GPSTAGS.get(key, key)
                    gps_tags[name] = gps_info[key]
                if ('GPSLatitude' in gps_tags and 'GPSLongitude' in gps_tags
                    and 'GPSLatitudeRef' in gps_tags and 'GPSLongitudeRef' in gps_tags):
                    lat = convert_gps(gps_tags['GPSLatitude'], gps_tags['GPSLatitudeRef'])
                    lon = convert_gps(gps_tags['GPSLongitude'], gps_tags['GPSLongitudeRef'])
                    metadata["GPS"] = {
                        "Latitude": lat,
                        "Longitude": lon,
                        "GoogleMaps": gps_to_link(lat, lon)
                    }
        else:
            metadata["info"] = "No EXIF metadata found in image"
    except Exception as e:
        metadata["error"] = f"Image metadata error: {e}"
    return metadata

def extract_pdf_metadata(path):
    try:
        with open(path, "rb") as f:
            reader = PyPDF2.PdfReader(f)
            doc_info = reader.metadata
            return {key[1:]: str(value) for key, value in doc_info.items()} if doc_info else {"info": "No metadata found in PDF"}
    except Exception as e:
        return {"error": f"PDF metadata error: {e}"}

def extract_docx_metadata(path):
    try:
        doc = docx.Document(path)
        props = doc.core_properties
        metadata = {}
        for attr in dir(props):
            if not attr.startswith("_") and not callable(getattr(props, attr)):
                value = getattr(props, attr)
                if value:
                    metadata[attr] = str(value)
        return metadata if metadata else {"info": "No metadata found in Word document"}
    except Exception as e:
        return {"error": f"DOCX metadata error: {e}"}

def extract_xlsx_metadata(path):
    try:
        wb = openpyxl.load_workbook(path, read_only=True)
        props = wb.properties
        return {
            "Title": props.title,
            "Subject": props.subject,
            "Creator": props.creator,
            "Created": str(props.created),
            "Modified": str(props.modified),
            "Category": props.category,
            "Keywords": props.keywords
        }
    except Exception as e:
        return {"error": f"Excel metadata error: {e}"}

def extract_text_metadata(path):
    try:
        with open(path, "r", encoding="utf-8") as f:
            lines = [f.readline().strip() for _ in range(10) if f.readline()]
        return {"Preview": lines if lines else "Empty text file"}
    except Exception as e:
        return {"error": f"Text metadata error: {e}"}

def extract_video_metadata(path):
    try:
        media_info = MediaInfo.parse(path)
        for track in media_info.tracks:
            if track.track_type == "General":
                return {
                    "Format": track.format,
                    "Duration (s)": round(track.duration / 1000, 2) if track.duration else None,
                    "File size (bytes)": track.file_size,
                    "Overall bit rate (bps)": track.overall_bit_rate,
                    "Encoded date": track.encoded_date,
                    "Tagged date": track.tagged_date
                }
        return {"info": "No video metadata found"}
    except Exception as e:
        return {"error": f"Video metadata error: {e}"}

def get_file_system_dates(path):
    try:
        stat = os.stat(path)
        created = datetime.fromtimestamp(stat.st_ctime)
        modified = datetime.fromtimestamp(stat.st_mtime)
        return {
            "Created": str(created),
            "Modified": str(modified),
            "Time Delta": str(modified - created)
        }
    except Exception as e:
        return {"error": f"File system error: {e}"}

# ----------------- Flask Routes ----------------- #
@app.route("/")
def index():
    return render_template("index.html")

@app.route("/analyze", methods=["POST"])
def analyze():
    file = request.files.get("file")
    if not file:
        return jsonify({"error": "No file uploaded"}), 400

    filepath = os.path.join("uploads", file.filename)
    os.makedirs("uploads", exist_ok=True)
    file.save(filepath)

    ext = os.path.splitext(filepath)[1].lower()
    metadata = {}

    if ext in [".jpg", ".jpeg", ".png"]:
        metadata["Image Metadata"] = extract_image_metadata(filepath)
    elif ext == ".pdf":
        metadata["PDF Metadata"] = extract_pdf_metadata(filepath)
    elif ext == ".docx":
        metadata["DOCX Metadata"] = extract_docx_metadata(filepath)
    elif ext == ".xlsx":
        metadata["XLSX Metadata"] = extract_xlsx_metadata(filepath)
    elif ext == ".txt":
        metadata["Text Metadata"] = extract_text_metadata(filepath)
    elif ext in [".mp4", ".avi"]:
        metadata["Video Metadata"] = extract_video_metadata(filepath)
    else:
        metadata["info"] = "Unsupported file type"

    metadata["File System"] = get_file_system_dates(filepath)
    return jsonify(metadata)

if __name__ == "__main__":
    app.run(debug=True)
