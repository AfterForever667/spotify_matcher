# spotify_matcher

This Python script is designed to bridge the gap between a personal music collection managed in a CSV file and the vast library of Spotify. It intelligently searches for each track from your CSV on Spotify, calculates a "confidence score" to determine the quality of the match, and outputs the results into a detailed, pre-formatted Excel file.

The primary goal is to overcome the common pitfalls of automated matching, such as finding incorrect live versions, remixes, or tracks by different artists with similar names.

## Key Features

-   **Intelligent Confidence Scoring:** Each potential match is assigned a score from 0-100 based on a weighted comparison of track name, artist, album, duration, and release year.
-   **Configurable Logic:** All scoring weights, penalties, and bonuses are controlled via an external `config.json` file, allowing you to fine-tune the matching algorithm without touching the code.
-   **Multiple Search Strategies:** For each track, the script employs several search queries—from highly specific to more general "sanitized" searches—to maximize the chances of finding the correct track, even with minor metadata discrepancies.
-   **Detailed Logging:** The output includes a "Details" sheet that logs every potential Spotify track considered, with a full breakdown of how its confidence score was calculated. This provides complete transparency into the matching process.
-   **Automated Formatting:** The final Excel output is automatically formatted with column filters, frozen panes, and auto-adjusted column widths for immediate analysis.
-   **Robust Error Handling:** The script is designed to handle missing optional data in the input CSV and gracefully reports any rows that could not be processed.
-   **Command-Line Interface:** Flexible execution with mandatory input/output file paths.

## Setup and Installation

### Prerequisites

-   Python 3.6 or higher.

### Step 1: Clone or Download the Repository

Download the files from this repository (`spotify_matcher.py`, `config.json`) and place them in a new folder on your computer.

### Step 2: Install Dependencies

Open a terminal or command prompt, navigate to the project folder, and run the following command to install the required Python libraries:

```bash
pip install pandas spotipy "fuzzywuzzy[speedup]" openpyxl
```

**Note for macOS/Linux users:** The quotes around `"fuzzywuzzy[speedup]"` are important to prevent your shell from misinterpreting the square brackets.

### Step 3: Get Spotify API Credentials

You need API keys from Spotify to allow the script to access its catalog.

1.  Go to the [Spotify Developer Dashboard](https://developer.spotify.com/dashboard/) and log in with your Spotify account.
2.  Click the **"Create App"** button.
3.  Give your app a name (e.g., "My Music Matcher") and a short description. Check the boxes to agree to the terms.
4.  Once the app is created, you will be on its dashboard. You will see your **Client ID**.
5.  Click **"Show client secret"** to reveal your **Client Secret**.
6.  You will need both of these long strings of characters for the next step.

### Step 4: Configure the Script

1.  Open the `spotify_matcher.py` file in a text editor.
2.  Near the top of the file, find the following lines:
    ```python
    # Paste your Spotify credentials directly here.
    CLIENT_ID = "YOUR_CLIENT_ID_HERE"
    SECRET_KEY = "YOUR_CLIENT_SECRET_HERE"
    ```
3.  Replace `"YOUR_CLIENT_ID_HERE"` and `"YOUR_CLIENT_SECRET_HERE"` with your actual credentials from the Spotify Developer Dashboard. **Keep the quotes.**

## Usage

### 1. Prepare Your Input CSV File

-   Your input data must be in a `.csv` file.
-   The CSV file **must** contain the following columns:
    -   `Artist`
    -   `Name` (Track Title)
    -   `Album`
    -   `Year`
    -   `Duration` (in M:SS or H:MM:SS format)
-   The script can also use these **optional** columns for higher accuracy:
    -   `Album Artist`
    -   `Track #`
    -   `Disc #`
-   Any superfluous columns (e.g., `Country`, `Genre`) will be ignored during matching and preserved in the final output.

### 2. (Optional) Customize the Matching Logic

You can edit the `config.json` file to change how the confidence score is calculated.

### 3. Run the Script from the Command Line

Open your terminal or command prompt and navigate to the folder containing the script. The script is run using the following format:

```bash
python spotify_matcher.py -input_csv <path_to_your_input_file> -output_excel <your_desired_output_name>
```

**Arguments:**

-   `-input_csv`: **(Required)** The path to your input CSV file. The `.csv` extension is optional.
-   `-output_excel`: **(Required)** The name for your output Excel file. The `.xlsx` extension is optional.
-   `-h` or `--help`: Display the help message with usage instructions.

**Important Note on File Paths:** The output Excel file will **always be created in the same directory where the `spotify_matcher.py` script is located**, even if your input CSV is in a different folder.

**Example:**

If your music file is on your Desktop, you might run:
```bash
python spotify_matcher.py -input_csv "/Users/YourUser/Desktop/my_music_collection.csv" -output_excel "spotify_results"
```
This will create a file named `spotify_results.xlsx` in your script's directory.

## Understanding the Output

The generated Excel file (`spotify_matches_output.xlsx`) contains two sheets:

### Summary Sheet

This sheet provides one row for each track in your original CSV file, preserving all original data and adding the following columns:
-   `Found on Spotify`: `True` or `False`.
-   `Include in Playlist`: **(New)** Defaults to the same value as `Found on Spotify`. You can manually change this to `TRUE` or `FALSE` to curate your final list before creating a playlist.
-   `Confidence`: The score of the best match found.
-   `Spotify Track ID`, `Spotify Name`, `Spotify Artist`, `Spotify Album`, `Spotify URL`: Details of the matched track if one was found.

This sheet is sorted alphabetically by **Album Artist**.

### Details Sheet

This sheet provides a transparent log of the script's decision-making process. It contains a row for *every potential Spotify match* that the script evaluated for all of your tracks. It includes columns such as:
-   `Local Artist`, `Local Track`, `Local Album`
-   `Match Found`: `True` or `False`, indicating the final outcome for this local track.
-   `Spotify Artist`, `Spotify Track`, `Spotify Album`
-   `Final Score`: The confidence score for this specific candidate.
-   ...and many more columns breaking down the score (`artist_similarity_%`, `year_penalty`, `album_artist_bonus`, etc.).

This sheet is sorted alphabetically by **Local Artist**.
