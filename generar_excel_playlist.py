import spotipy
from spotipy.oauth2 import SpotifyClientCredentials
import pandas as pd
import os
import requests
from io import BytesIO
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Alignment, Font, PatternFill
from PIL import Image as PILImage

# Paleta de colores de Spotify
SPOTIFY_GREEN = "1DB954"
DARK_GRAY = "191414"
LIGHT_GRAY = "EFEFEF"

# Carpeta donde guardar
OUTPUT_FOLDER = "PlaylistsExcel"
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

def main():
    client_id = "71114e96572a4b759750f90f89653e12"
    client_secret = "44374ebc9731491e87bee7fad0156a2c"

    auth_manager = SpotifyClientCredentials(client_id=client_id, client_secret=client_secret)
    sp = spotipy.Spotify(auth_manager=auth_manager)

    playlist_url = input("🎵 Pega la URL de la playlist de Spotify: ")
    playlist_id = playlist_url.split("/")[-1].split("?")[0]

    playlist = sp.playlist(playlist_id)
    playlist_name = playlist['name']
    playlist_image_url = playlist['images'][0]['url']

    # Descargar imagen
    img_data = requests.get(playlist_image_url).content
    img_path = os.path.join(OUTPUT_FOLDER, f"{playlist_name}_image.png")
    with open(img_path, 'wb') as f:
        f.write(img_data)

    tracks_data = []
    for item in playlist['tracks']['items']:
        track = item['track']
        name = track['name']
        url = track['external_urls']['spotify']
        artist = ", ".join([a['name'] for a in track['artists']])
        duration_ms = track['duration_ms']
        minutes = duration_ms // 60000
        seconds = (duration_ms % 60000) // 1000
        duration = f"{minutes}:{seconds:02d}"
        tracks_data.append((name, artist, duration, url))

    # Crear Excel con openpyxl
    wb = Workbook()
    ws = wb.active
    ws.title = "Playlist"

    # Títulos
    headers = ["Canción", "Artista(s)", "Duración"]
    ws.append(headers)

    # Estilos
    header_font = Font(bold=True, color="FFFFFF", name="Arial")
    header_fill = PatternFill("solid", fgColor=SPOTIFY_GREEN)
    center_align = Alignment(horizontal="center", vertical="center")
    default_font = Font(name="Arial", size=11)

    for col in range(1, 4):
        cell = ws.cell(row=1, column=col)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center_align

    # Insertar datos con hipervínculo en nombre
    for row_idx, (name, artist, duration, url) in enumerate(tracks_data, start=2):
        ws.cell(row=row_idx, column=1).value = name
        ws.cell(row=row_idx, column=1).hyperlink = url
        ws.cell(row=row_idx, column=1).font = Font(name="Arial", color="0563C1", underline="single")

        ws.cell(row=row_idx, column=2).value = artist
        ws.cell(row=row_idx, column=2).font = default_font

        ws.cell(row=row_idx, column=3).value = duration
        ws.cell(row=row_idx, column=3).font = default_font
        ws.cell(row=row_idx, column=3).alignment = center_align

    # Ajustar ancho de columnas a 257px ≈ 38.5 unidades de Excel
    ws.column_dimensions["A"].width = 38.5
    ws.column_dimensions["B"].width = 38.5
    ws.column_dimensions["C"].width = 12

    # Insertar imagen de la playlist
    img = PILImage.open(img_path)
    img = img.resize((150, 150))
    img.save(img_path)

    xl_img = XLImage(img_path)
    xl_img.anchor = "E2"
    ws.add_image(xl_img)

    # Insertar solo el nombre de la playlist
    name_cell = ws.cell(row=10, column=5)
    name_cell.value = f"Playlist: {playlist_name}"
    name_cell.font = Font(name="Arial", size=14, bold=True)
    name_cell.alignment = Alignment(horizontal="center", vertical="center")

    # Guardar archivo
    safe_name = "".join(c if c.isalnum() or c in " _-" else "_" for c in playlist_name)
    excel_file = os.path.join(OUTPUT_FOLDER, f"{safe_name}.xlsx")
    wb.save(excel_file)
    print(f"\n✅ Excel generado: {excel_file}")

if __name__ == "__main__":
    main()
