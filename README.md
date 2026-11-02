
# 🌊 BluePath
## Blue Economy × Lean Construction Decision Support Platform

### About BluePath

BluePath is an interactive decision-support application designed to explore how Lean principles can facilitate the implementation of Blue Economy practices in marine construction projects.

The application provides a structured approach to identifying marine construction impacts, exploring suitable Blue Economy practices, recognising implementation challenges, and identifying relevant Lean principles to address those challenges.

### Purpose

The primary purpose of BluePath is to support decision-making towards sustainable marine construction by connecting:

1. Impacts of Marine Construction
2. Blue Economy Practices
3. Challenges to Implement Blue Economy Practices
4. Lean Principles

### How to Use the Application

**Step 1: Welcome**

Enter your name and select Start Exploring BluePath.

**Step 2: Identify Marine Construction Impacts**

Explore the available impacts and select the impact that you wish to manage.

**Step 3: Explore Suitable Blue Economy Practices**

Review the Blue Economy practices associated with the selected marine construction impact.

**Step 4: Identify Implementation Challenges**

Select a Blue Economy practice to explore the challenges associated with its implementation in marine construction.

**Step 5: Explore Relevant Lean Principles**

Select a particular challenge to identify the corresponding Lean principles that may help minimise that challenge.

### Application Navigation

Users can navigate between the different stages using the Back and Home options.

The application also provides an option to start a new decision-making journey.

### Database

The application retrieves information from an Excel database named `portal.xlsx`.

The Excel database must contain the following four columns:

- Impact of Marine Construction
- Blue Economy Practice
- Challenges to Implement Blue Economy Practice
- Lean Principles

The Excel file must be stored in the same repository as `app.py`.

### Technologies Used

- Python
- Streamlit
- Pandas
- OpenPyXL
- GitHub
- Streamlit Community Cloud

### Repository Structure

```text
BluePath/
│
├── app.py
├── portal.xlsx
├── README.md
└── requirements.txt
```

### Running the Application Locally

1. Clone or download the GitHub repository.
2. Install the required Python packages.
3. Open the repository directory.
4. Run the following command:

```bash
streamlit run app.py
```

### Deployment

The application can be deployed using Streamlit Community Cloud by connecting the GitHub repository and selecting `app.py` as the main application file.

### Important Note

BluePath is intended to provide structured decision-support information based on the relationships recorded in its Excel database. The identified practices and Lean principles should be evaluated in relation to the requirements and conditions of individual marine construction projects.

---

**BluePath | Blue Economy × Lean Construction**

Supporting sustainable and responsible marine construction.
