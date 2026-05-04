# ⚾ Tunneling+

A comprehensive Streamlit app for analyzing baseball tunneling metrics using Statcast data. Built on the V18 pipeline with advanced pitch tunneling calculations.

[![Open in Streamlit](https://static.streamlit.io/badges/streamlit_badge_black_white.svg)](https://tunneling-plus.streamlit.app/)

## Features

- 📊 **Interactive Leaderboard**: View tunneling rankings for qualified pitchers
- 🃏 **Player Cards**: Generate beautiful JPG cards with pitch tunneling analysis
- 📥 **Data Export**: Download leaderboards as Excel files
- 🔬 **Diagnostic Tools**: Explore raw Statcast data and calculations
- 🎯 **Real-time Updates**: Fresh data loading with caching

## Model Details

- **Pipeline**: V18 (Statcast Edition)
- **Data Source**: MLB Statcast via pybaseball
- **Metrics**: Tunneling+, Tunnel Pairs, Speed Pairs, Temporal Analysis
- **Filtering**: All pitches included (no windup filter)

## Local Development

### Prerequisites

- Python 3.8+
- Virtual environment recommended

### Setup

1. **Clone the repository**
   ```bash
   git clone https://github.com/robmcolonna123/Tunn.git
   cd Tunn
   ```

2. **Create virtual environment**
   ```bash
   python -m venv .venv
   source .venv/bin/activate  # On Windows: .venv\Scripts\activate
   ```

3. **Install dependencies**
   ```bash
   pip install -r requirements.txt
   ```

4. **Install Playwright for JPG rendering** (optional)
   ```bash
   pip install playwright
   playwright install chromium
   ```

5. **Run the app**
   ```bash
   streamlit run app.py
   ```

## Deployment

### Streamlit Cloud

1. **Fork this repository** to your GitHub account
2. **Go to [share.streamlit.io](https://share.streamlit.io)**
3. **Connect your GitHub account** and select this repository
4. **Deploy**: The app will automatically deploy with the `app.py` entry point

### Other Platforms

The app can also be deployed to:
- **Heroku**: Use the `Procfile` and build configuration
- **Railway**: Automatic deployment from GitHub
- **Vercel**: For static deployments (limited functionality)

## Data Sources

- **MLB Statcast**: Pitch-by-pitch data via pybaseball
- **Real-time Updates**: Data cached for 1 hour to balance freshness and performance
- **Season Coverage**: 2021-2026 seasons supported

## Architecture

- **Frontend**: Streamlit web interface
- **Backend**: Python data processing pipeline
- **Visualization**: HTML/CSS/JS rendered via Playwright for JPG export
- **Caching**: Streamlit's built-in caching for performance

## Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Test locally
5. Submit a pull request

## License

See LICENSE file for details.

## Credits

Built by Robert Colonna - Advanced baseball analytics and tunneling research.
