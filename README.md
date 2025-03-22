# OKR Scoring System

![OKR Scoring System](https://img.shields.io/badge/Status-Active-success)
![Streamlit](https://img.shields.io/badge/Built%20with-Streamlit-FF4B4B)

A comprehensive web application for tracking, calculating, and visualizing OKR (Objectives and Key Results) scores within an organization. This tool helps teams monitor their performance metrics and provides insights into user engagement with the OKR process.

**Live Demo**: [OKR Scoring System](https://okr-scroring-aplus.streamlit.app/)

## 📋 Features

- **Real-time OKR Data Integration**: Connects with Base Goal API to fetch real-time OKR data
- **Automatic Score Calculation**: Calculates user scores based on multiple criteria:
  - OKR completion status
  - Weekly check-in compliance
  - OKR movement (progress compared to previous month)
- **Interactive Dashboard**: Visualizes key metrics and user performance data
- **Excel Export**: Generates formatted Excel reports with detailed scoring breakdowns
- **Historical Data Tracking**: Stores and retrieves historical OKR data via Google Sheets integration

## 🚀 Getting Started

### Prerequisites

- Python 3.7+
- Streamlit
- Access to Base Goal API
- Google Sheets API setup for historical data

### Installation

1. Clone the repository
```bash
git clone https://github.com/yourusername/okr-scoring-system.git
cd okr-scoring-system
```

2. Install dependencies
```bash
pip install -r requirements.txt
```

3. Set up environment variables
```bash
# Create a .env file with the following variables
GOAL_ACCESS_TOKEN="your_goal_access_token"
ACCOUNT_ACCESS_TOKEN="your_account_access_token"
GOOGLE_SHEETS_API_URL="your_google_sheets_api_url"
```

4. Run the application
```bash
streamlit run app.py
```

## 🔧 How It Works

The OKR Scoring System evaluates user performance based on three main criteria:

1. **OKR Setup** (1 point)
   - User has individual OKRs set up in Base Goal

2. **Check-ins** (0.5 point)
   - Weekly check-ins (at least 3 per month)

3. **OKR Movement** (0.15 - 2.5 points)
   - < 10%: 0.15 points
   - 10-25%: 0.25 points
   - 26-30%: 0.5 points
   - 31-50%: 0.75 points
   - 51-80%: 1.25 points
   - 81-99%: 1.5 points
   - 100% or breakthrough: 2.5 points

The system automatically fetches data, calculates scores, and provides visual reporting through an intuitive interface.

## 🏗️ Project Structure

```
okr-scoring-system/
├── app.py                # Main Streamlit application
├── requirements.txt      # Dependencies
├── .env                  # Environment variables (git-ignored)
└── README.md             # Project documentation
```

## 📈 Data Flow

1. User selects a quarterly cycle from the dropdown
2. System fetches data from Base Goal API (user accounts, check-ins, KRs, cycle data)
3. Calculates scores based on predefined criteria
4. Displays results in an interactive dashboard
5. Allows export to formatted Excel report

## 🔒 API Integration

The application integrates with:
- Base Goal API for real-time OKR data
- Base Account API for user information
- Google Sheets API for historical data storage

## 🛠️ Customization

You can modify the scoring criteria by adjusting the parameters in the `calculate_score` method of the `User` class.
## 👥 Contributors

- [Tran Thanh Son](https://github.com/FOX2920)
- [Vo Le Thanh Phat](https://github.com/F4tt)

## 🙏 Acknowledgements

- [Base.vn](https://base.vn) for their Goal API
- [Streamlit](https://streamlit.io) for the web application framework
