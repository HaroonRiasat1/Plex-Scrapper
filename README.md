# PLEXScrapper

PLEXScrapper is a Python-based tool that scrapes movie information from a Plex Media Server. It takes an Excel file as input, which contains a list of movie titles, and outputs the corresponding metadata and details of the movies available on your Plex server.

## Features

- **Excel Integration**: Takes an Excel file as input to define the list of movies to scrape.
- **Plex API**: Connects to your Plex Media Server using the Plex API to retrieve movie metadata.
- **Detailed Output**: Outputs detailed information for each movie, including title, year, genre, director, summary, and more.
- **Flexible Usage**: Easily adaptable for different use cases by modifying the input Excel file.

## Requirements

- Python 3.x
- Plex Media Server with API access enabled
- Required Python packages:
  - `pandas`
  - `requests`
  - `openpyxl`
  - `plexapi`

## Installation

1. **Clone the Repository**:

   ```bash
   git clone https://github.com/yourusername/PLEXScrapper.git
   cd PLEXScrapper
   ```

2. **Install Required Packages**:

   Install the necessary Python packages using `pip`:

   ```bash
   pip install -r requirements.txt
   ```

3. **Plex API Token**:

   To connect to your Plex server, you'll need a Plex API token. You can find instructions on how to retrieve your Plex token [here](https://support.plex.tv/articles/204059436-finding-an-authentication-token-x-plex-token/).

## Usage

1. **Prepare the Input Excel File**:

   The input Excel file should contain a list of movie titles you want to scrape. Ensure the file is saved in `.xlsx` format, with the movie titles listed in a column named `Title`.

2. **Run the Scrapper**:

   Execute the script with the following command:

   ```bash
   python plexscrapper.py --input <path_to_input_excel> --output <path_to_output_excel> --token <plex_api_token> --url <plex_server_url>
   ```

   - `<path_to_input_excel>`: Path to the input Excel file containing movie titles.
   - `<path_to_output_excel>`: Path where the output Excel file will be saved.
   - `<plex_api_token>`: Your Plex API token.
   - `<plex_server_url>`: The URL of your Plex server.

3. **Example Command**:

   ```bash
   python plexscrapper.py --input movies.xlsx --output output.xlsx --token YOUR_PLEX_TOKEN --url http://localhost:32400
   ```

   This command will scrape the movies listed in `movies.xlsx` from the Plex server running at `http://localhost:32400` and save the output in `output.xlsx`.

## Output

The output Excel file will contain the following columns:

- **Title**: The title of the movie.
- **Year**: The release year of the movie.
- **Genre**: The genres associated with the movie.
- **Director**: The director(s) of the movie.
- **Summary**: A brief summary or plot of the movie.
- **Rating**: The rating of the movie (if available).
- **Duration**: The duration of the movie in minutes.
- **Added Date**: The date the movie was added to the Plex library.

## Troubleshooting

- **Invalid Token**: Ensure that your Plex API token is valid and correctly entered.
- **Server URL**: Double-check that your Plex server URL is accessible and correct.
- **Excel Format**: Ensure that your input Excel file is in `.xlsx` format and the column name is `Title`.

## Contributing

If you would like to contribute to PLEXScrapper, please fork the repository and submit a pull request. We welcome all contributions!

## License

PLEXScrapper is licensed under the MIT License. See the [LICENSE](LICENSE) file for more details.

## Acknowledgments

- [PlexAPI](https://github.com/pkkid/python-plexapi) - Python Plex API used for interacting with Plex Media Server.
- [pandas](https://pandas.pydata.org/) - Python library for data manipulation and analysis.
- [openpyxl](https://openpyxl.readthedocs.io/en/stable/) - A Python library to read/write Excel 2010 xlsx/xlsm/xltx/xltm files.

---

For any issues or support, please open an issue on the GitHub repository.

Happy Scraping!
