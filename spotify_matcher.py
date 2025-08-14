"""
Spotify Matcher

This script reads a CSV file containing music track information, searches for each track on
Spotify using its Web API, and calculates a confidence score to find the best match.
The results, including the Spotify track ID, URL, and a detailed scoring breakdown,
are saved to a multi-sheet, formatted Excel file.

The script is executed via the command line, requiring paths for the input CSV
and the desired output Excel file.
"""

# --- Core Libraries ---
import os
import re
import json
import time
import argparse

# --- Third-Party Libraries ---
import pandas as pd
import spotipy
from spotipy.oauth2 import SpotifyClientCredentials
from fuzzywuzzy import fuzz
from openpyxl.utils import get_column_letter

# --- User Configuration ---
# IMPORTANT: Paste your Spotify API credentials directly here.
# You can get these from the Spotify Developer Dashboard: https://developer.spotify.com/dashboard/
CLIENT_ID = "YOUR_CLIENT_ID_HERE"
SECRET_KEY = "YOUR_CLIENT_SECRET_HERE"

# --- Helper Functions ---

def apply_formatting(worksheet, dataframe):
    """Applies standard formatting to an openpyxl worksheet for better readability."""
    if dataframe.empty: return
    for i, column_name in enumerate(dataframe.columns):
        column_len = dataframe[column_name].astype(str).map(len).max()
        header_len = len(str(column_name))
        adjusted_width = max(column_len, header_len) + 2
        worksheet.column_dimensions[get_column_letter(i + 1)].width = adjusted_width
    worksheet.freeze_panes = 'E2'
    worksheet.auto_filter.ref = worksheet.dimensions

def load_config(config_path='config.json'):
    """Loads the scoring logic and thresholds from an external JSON configuration file."""
    print(f"Loading configuration from {config_path}...")
    with open(config_path, 'r') as f:
        config = json.load(f)
    return config

def setup_spotipy():
    """Initializes and returns an authenticated Spotipy client instance."""
    if not CLIENT_ID or not SECRET_KEY or CLIENT_ID == "YOUR_CLIENT_ID_HERE":
        raise ValueError("Spotify API credentials not set. Please edit the CLIENT_ID and SECRET_KEY constants at the top of the script.")
    client_credentials_manager = SpotifyClientCredentials(client_id=CLIENT_ID, client_secret=SECRET_KEY)
    sp = spotipy.Spotify(client_credentials_manager=client_credentials_manager)
    return sp

def clean_string(text):
    """Removes common extra info like "(Remastered)" or "[Instrumental]" for better matching."""
    if not isinstance(text, str): return ""
    text = re.sub(r'\[.*?\]', '', text)
    text = re.sub(r'\(.*?\)', '', text)
    return text.strip()

def sanitize_for_search(text):
    """Cleans a string more aggressively for searching, removing roman numerals and versioning."""
    if not isinstance(text, str): return ""
    sanitized_text = clean_string(text)
    roman_numerals = [r'\bV\b', r'\bIV\b', r'\bIII\b', r'\bII\b', r'\bI\b']
    for pattern in roman_numerals:
        sanitized_text = re.sub(pattern, '', sanitized_text, flags=re.IGNORECASE)
    sanitized_text = re.sub(r'\s+', ' ', sanitized_text).strip()
    return sanitized_text

def convert_duration_to_ms(duration_str):
    """Converts a 'M:SS' or 'H:MM:SS' duration string to milliseconds."""
    if not isinstance(duration_str, str): return 0
    parts = list(map(int, str(duration_str).split(':')))
    if len(parts) == 3: return (parts[0] * 3600 + parts[1] * 60 + parts[2]) * 1000
    if len(parts) == 2: return (parts[0] * 60 + parts[1]) * 1000
    return 0

# --- Core Logic ---

def calculate_confidence(local_track, spotify_track, config):
    """Calculates a confidence score (0-100) based on how well local and Spotify data match."""
    breakdown = {}
    local_clean_artist, spotify_clean_artist = clean_string(local_track['Artist']), clean_string(spotify_track['artist_name'])
    local_clean_name, spotify_clean_name = clean_string(local_track['Name']), clean_string(spotify_track['track_name'])

    artist_similarity = fuzz.token_set_ratio(local_clean_artist, spotify_clean_artist)
    breakdown['artist_similarity_%'], breakdown['artist_score'] = artist_similarity, (artist_similarity / 100) * config['base_weights']['artist']
    if artist_similarity < config['rules']['min_artist_similarity']:
        breakdown['final_score'], breakdown['reason_for_zero'] = 0, f"Artist similarity {artist_similarity}% is below threshold {config['rules']['min_artist_similarity']}%"
        return 0, breakdown

    track_similarity = fuzz.token_set_ratio(local_clean_name, spotify_clean_name)
    breakdown['track_similarity_%'], breakdown['track_name_score'] = track_similarity, (track_similarity / 100) * config['base_weights']['track_name']
    
    if artist_similarity == 100 and set(local_clean_artist.lower().split()) != set(spotify_clean_artist.lower().split()):
        breakdown['artist_unmatched_words_penalty'] = config['penalties']['unmatched_words_penalty']
    if track_similarity == 100 and set(local_clean_name.lower().split()) != set(spotify_clean_name.lower().split()):
        breakdown['track_unmatched_words_penalty'] = config['penalties']['unmatched_words_penalty']
    if artist_similarity == 100 and track_similarity == 100:
        breakdown['perfect_core_bonus'] = config['bonuses'].get('perfect_core_match', 0)

    album_similarity = fuzz.token_set_ratio(local_track['Album'], spotify_track['album_name'])
    breakdown['album_similarity_%'] = album_similarity
    if album_similarity > 90: breakdown['album_name_bonus'] = config['bonuses']['strong_album_match']
    elif album_similarity < 50: breakdown['album_name_penalty'] = config['penalties']['album_mismatch']

    local_album_artist, spotify_album_artist = local_track.get('Album Artist', local_track['Artist']), spotify_track.get('album_artist_name', spotify_track['artist_name'])
    is_local_va, is_spotify_va = 'various artists' in str(local_album_artist).lower(), 'various artists' in str(spotify_album_artist).lower()
    if is_local_va and is_spotify_va: breakdown['album_artist_bonus'] = config['bonuses']['album_artist_match']
    elif fuzz.ratio(local_album_artist, spotify_album_artist) > 90: breakdown['album_artist_bonus'] = config['bonuses']['album_artist_match']
    elif is_local_va != is_spotify_va and artist_similarity < 100: breakdown['album_artist_penalty'] = config['penalties']['album_artist_mismatch']

    if album_similarity > 85:
        local_track_num, spotify_track_num = local_track.get('Track #'), spotify_track.get('track_number')
        local_disc_num, spotify_disc_num = local_track.get('Disc #', 1), spotify_track.get('disc_number', 1)
        try:
            if local_track_num and pd.notna(local_track_num) and int(local_disc_num) == int(spotify_disc_num) and int(local_track_num) == int(spotify_track_num):
                breakdown['track_number_bonus'] = config['bonuses']['track_number_match']
            elif local_track_num and pd.notna(local_track_num):
                breakdown['track_number_penalty'] = config['penalties']['track_number_mismatch']
        except (ValueError, TypeError): pass

    duration_diff = abs(local_track['duration_ms'] - spotify_track['duration_ms'])
    duration_diff_percent = (duration_diff / local_track['duration_ms']) * 100 if local_track['duration_ms'] > 0 else 100
    breakdown['duration_diff_%'] = round(duration_diff_percent, 2)
    if duration_diff_percent > config['rules']['duration_diff_large_percent']: breakdown['duration_penalty'] = config['penalties']['duration_diff_large']
    elif duration_diff_percent > config['rules']['duration_diff_medium_percent']: breakdown['duration_penalty'] = config['penalties']['duration_diff_medium']

    year_diff = abs(local_track['Year'] - int(str(spotify_track['release_year'])[:4]))
    breakdown['year_difference'] = year_diff
    if year_diff > config['rules']['year_diff_large_years']: breakdown['year_penalty'] = config['penalties']['year_diff_large']
    elif year_diff > config['rules']['year_diff_medium_years']: breakdown['year_penalty'] = config['penalties']['year_diff_medium']

    local_is_live = 'live' in local_track['Name'].lower() or 'live' in local_track['Album'].lower()
    spotify_is_live = 'live' in spotify_track['track_name'].lower() or 'live' in spotify_track['album_name'].lower()
    if local_is_live != spotify_is_live: breakdown['live_mismatch_penalty'] = config['penalties']['live_mismatch']

    total_score = sum(v for k, v in breakdown.items() if 'score' in k or 'penalty' in k or 'bonus' in k)
    final_score = max(0, min(100, int(total_score)))
    breakdown['final_score'] = final_score
    return final_score, breakdown

def search_for_track(sp, local_track, config):
    """Performs a series of searches on Spotify to find all possible candidates for a track."""
    best_match, highest_confidence, detailed_logs = (None, 0, [])
    clean_artist_name = clean_string(local_track['Artist'])
    original_clean_track_name, sanitized_for_search_name = clean_string(local_track['Name']), sanitize_for_search(local_track['Name'])
    search_queries = [
        f"track:\"{original_clean_track_name}\" artist:\"{clean_artist_name}\" album:\"{clean_string(local_track['Album'])}\"",
        f"track:\"{original_clean_track_name}\" artist:\"{clean_artist_name}\" year:{local_track['Year']}",
        f"track:\"{sanitized_for_search_name}\" artist:\"{clean_artist_name}\"",
        f"track:\"{original_clean_track_name}\" artist:\"{clean_artist_name}\"",
        f"{local_track['Name']} {local_track['Artist']}"
    ]
    processed_spotify_ids = set()
    for query in search_queries:
        try:
            results = sp.search(q=query, type='track', limit=10)
            if not results or not results['tracks']['items']: continue
            for item in results['tracks']['items']:
                if item['id'] in processed_spotify_ids: continue
                processed_spotify_ids.add(item['id'])
                spotify_track = {'track_name': item['name'], 'artist_name': ', '.join(a['name'] for a in item['artists']), 'album_name': item['album']['name'], 'album_artist_name': ', '.join(a['name'] for a in item['album']['artists']), 'release_year': item['album']['release_date'], 'duration_ms': item['duration_ms'], 'id': item['id'], 'url': item['external_urls']['spotify'], 'track_number': item.get('track_number'), 'disc_number': item.get('disc_number')}
                confidence, breakdown = calculate_confidence(local_track, spotify_track, config)
                log_entry = {'Local Artist': local_track['Artist'], 'Local Track': local_track['Name'], 'Local Album': local_track['Album'], 'Spotify Artist': spotify_track['artist_name'], 'Spotify Track': spotify_track['track_name'], 'Spotify Album': spotify_track['album_name'], 'Spotify Year': str(spotify_track['release_year'])[:4], 'Final Score': breakdown.get('final_score', 0), **breakdown}
                detailed_logs.append(log_entry)
                if confidence > highest_confidence: highest_confidence, best_match = confidence, spotify_track
            if highest_confidence > 95: break
        except Exception as e:
            print(f"An error occurred during search for query '{query}': {e}")
            time.sleep(2)
    return best_match, highest_confidence, detailed_logs

# --- Main Execution Block ---

def main(input_csv_path, output_excel_path):
    """Main function to orchestrate the entire process."""
    config = load_config()
    sp = setup_spotipy()
    df = pd.read_csv(input_csv_path)
    
    required_cols = ['Artist', 'Name', 'Album', 'Year', 'Duration']
    optional_cols = ['Album Artist', 'Track #', 'Disc #']
    for col in required_cols:
        if col not in df.columns: raise KeyError(f"Input CSV is missing required column: '{col}'")
    for col in optional_cols:
        if col not in df.columns: df[col] = None
    
    summary_results, all_detailed_logs, total_rows = [], [], len(df)
    print(f"\nStarting to process {total_rows} tracks...")
    
    for index, row in df.iterrows():
        summary_row = row.to_dict()
        try:
            if pd.isna(row['Year']): raise ValueError("Year is missing")
            local_track = {'Artist': row['Artist'], 'Name': row['Name'], 'Album': row['Album'], 'Album Artist': row['Album Artist'], 'Year': int(row['Year']), 'Duration': row['Duration'], 'duration_ms': convert_duration_to_ms(row['Duration']), 'Track #': row['Track #'], 'Disc #': row['Disc #']}
            print(f"\n--- Processing {index + 1}/{total_rows}: {row['Artist']} - {row['Name']} ---")
            best_match, confidence, detailed_logs = search_for_track(sp, local_track, config)
            final_match_found = bool(best_match and confidence >= config['confidence_threshold'])
            for log in detailed_logs: log['Match Found'] = final_match_found
            all_detailed_logs.extend(detailed_logs)
            
            # <<< MODIFICATION START >>>
            if final_match_found:
                print(f"✅ Found Match: '{best_match['track_name']}' with confidence: {confidence}%")
                summary_row.update({'Found on Spotify': True, 'Include in Playlist': True, 'Confidence': f"{confidence}%", 'Spotify Track ID': best_match['id'], 'Spotify Name': best_match['track_name'], 'Spotify Artist': best_match['artist_name'], 'Spotify Album': best_match['album_name'], 'Spotify URL': best_match['url']})
            else:
                print(f"❌ No suitable match found. Highest confidence: {confidence}%")
                summary_row.update({'Found on Spotify': False, 'Include in Playlist': False, 'Confidence': f"{confidence}% (Below Threshold)", 'Spotify Track ID': '', 'Spotify Name': '', 'Spotify Artist': '', 'Spotify Album': '', 'Spotify URL': ''})
            # <<< MODIFICATION END >>>
        except Exception as e:
            print(f"\n--- ERROR Processing {index + 1}/{total_rows}: {row.get('Artist', 'N/A')} - {row.get('Name', 'N/A')} ---")
            print(f"   Could not process row due to bad data: {e}. Marking as not found.")
            summary_row.update({'Found on Spotify': False, 'Include in Playlist': False, 'Confidence': "Error", 'Spotify Track ID': '', 'Spotify Name': '', 'Spotify Artist': '', 'Spotify Album': '', 'Spotify URL': ''})
        
        summary_results.append(summary_row)

    print("\nProcessing complete. Creating Excel file with two sheets...")
    summary_df = pd.DataFrame(summary_results)
    detailed_df = pd.DataFrame(all_detailed_logs)
    
    # Sort the DataFrames before writing to Excel
    summary_df = summary_df.sort_values(by='Album Artist', ascending=True, na_position='first')
    
    # <<< MODIFICATION START >>>
    # Reorder the summary columns to place 'Include in Playlist' after 'Found on Spotify'
    summary_cols = summary_df.columns.tolist()
    if 'Found on Spotify' in summary_cols:
        # Move 'Include in Playlist' to the correct position
        if 'Include in Playlist' in summary_cols:
            summary_cols.remove('Include in Playlist')
        insert_pos = summary_cols.index('Found on Spotify') + 1
        summary_cols.insert(insert_pos, 'Include in Playlist')
        summary_df = summary_df[summary_cols]
    # <<< MODIFICATION END >>>

    if not detailed_df.empty:
        first_cols = ['Local Artist', 'Local Track', 'Local Album', 'Match Found', 'Spotify Artist', 'Spotify Track', 'Spotify Album', 'Spotify Year', 'Final Score']
        other_cols = [col for col in detailed_df.columns if col not in first_cols]
        detailed_df = detailed_df[first_cols + other_cols]
        detailed_df = detailed_df.sort_values(by=['Local Artist', 'Local Track', 'Final Score'], ascending=[True, True, False])

    with pd.ExcelWriter(output_excel_path, engine='openpyxl') as writer:
        summary_df.to_excel(writer, sheet_name='Summary', index=False)
        if not detailed_df.empty:
            detailed_df.to_excel(writer, sheet_name='Details', index=False)
        
        summary_sheet = writer.sheets['Summary']
        print("Applying formatting to Summary sheet...")
        apply_formatting(summary_sheet, summary_df)
        
        if 'Details' in writer.sheets and not detailed_df.empty:
            details_sheet = writer.sheets['Details']
            print("Applying formatting to Details sheet...")
            apply_formatting(details_sheet, detailed_df)
        
    print(f"✨ Success! Output saved to {output_excel_path}")

if __name__ == '__main__':
    parser = argparse.ArgumentParser(description="Finds Spotify tracks that match a local music collection CSV.", formatter_class=argparse.RawTextHelpFormatter)
    parser.add_argument('-input_csv', required=True, metavar='INPUT_PATH', help="Path to the input CSV file. The .csv extension is optional.")
    parser.add_argument('-output_excel', required=True, metavar='OUTPUT_NAME', help="Name for the output Excel file. The .xlsx extension is optional.")
    args = parser.parse_args()
    input_file = args.input_csv if args.input_csv.lower().endswith('.csv') else args.input_csv + '.csv'
    output_file = args.output_excel if args.output_excel.lower().endswith('.xlsx') else args.output_excel + '.xlsx'
    main(input_file, output_file)