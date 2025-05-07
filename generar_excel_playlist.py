import spotipy
from spotipy.oauth2 import SpotifyClientCredentials
import pandas as pd
import re  # <- para limpiar el nombre del archivo

def main():
    client_id = "71114e96572a4b759750f90f89653e12"
    client_secret = "44374ebc9731491e87bee7fad0156a2c"

    auth_manager = SpotifyClientCredentials(client_id=client_id, client_secret=client_secret)
    sp = spotipy.Spotify(auth_manager=auth_manager)

    playlist_url = input("Pega la URL de la playlist de Spotify: ")
    playlist_id = playlist_url.split("/")[-1].split("?")[0]

    playlist = sp.playlist(playlist_id)
    playlist_name = playlist['name']

    # 🔧 Limpiar el nombre del archivo quitando caracteres no válidos
    playlist_name = re.sub(r'[\\/*?:"<>|]', "", playlist_name)

    tracks_data = []
    for item in playlist['tracks']['items']:
        track = item['track']
        name = track['name']
        artist = ", ".join([a['name'] for a in track['artists']])
        duration_ms = track['duration_ms']
        minutes = duration_ms // 60000
        seconds = (duration_ms % 60000) // 1000
        duration = f"{minutes}:{seconds:02d}"

        tracks_data.append({
            'Canción': name,
            'Artista(s)': artist,
            'Duración': duration
        })

    df = pd.DataFrame(tracks_data)
    excel_file = f"{playlist_name}.xlsx"
    df.to_excel(excel_file, index=False)
    print(f"\n✅ Excel generado: {excel_file}")

if __name__ == "__main__":
    main()
