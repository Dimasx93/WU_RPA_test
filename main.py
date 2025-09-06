import requests
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from typing import Dict, List, Any

# =====================================================
# Load Users from CSV
# =====================================================
# Users are stored in a CSV file called "ParaBank users.csv"
users: pd.DataFrame = pd.read_csv("ParaBank users.csv")

# Fields required for registration
REQUIRED_FIELDS: List[str] = [
    "First Name", "Last Name", "Address", "City", "State", "Zip Code",
    "Phone Number", "SSN", "Username", "Password"
]

# Ensure only relevant columns are kept
users = users[
    [
        "First Name", "Last Name", "Address", "City", "State", "Zip Code",
        "Phone Number", "SSN", "Username", "Password",
        "Initial Deposit", "DOB", "Debit Card", "CVV"
    ]
]

# =====================================================
# Get Live USD → EUR Exchange Rate
# =====================================================
try:
    resp: requests.Response = requests.get("https://open.er-api.com/v6/latest/USD", timeout=10)
    resp.raise_for_status()
    rate: float = resp.json().get("rates", {}).get("EUR")
    if rate is None:
        raise ValueError("EUR rate not found in API response")
    USD_TO_EUR: float = rate
    print(f"Live USD→EUR rate: 1 USD = {USD_TO_EUR:.2f} EUR")
except Exception as ex:
    print(f"⚠ Failed to fetch live rate, using fallback 0.92. Error: {ex}")
    USD_TO_EUR: float = 0.86  # Fallback

# =====================================================
# Loan Setup
# =====================================================
LOAN_AMOUNT_USD: int = 10000
report: List[Dict[str, Any]] = []  # Report data collected for all users

# =====================================================
# Process Each Customer
# =====================================================
for _, user in users.iterrows():
    username: str = str(user["Username"])
    password: str = str(user["Password"])

    # Copy user info into dictionary for reporting
    user_report: Dict[str, Any] = user.to_dict()

    # Initialize process tracking fields
    user_report.update({
        "Registration Status": None,
        "Login Status": None,
        "Account Opened": None,
        "Loan Requested": None,
        "Error": None
    })

    # -------------------------
    # Step 1: Validate required fields
    # -------------------------
    missing: List[str] = [
        f for f in REQUIRED_FIELDS if pd.isna(user[f]) or str(user[f]).strip() == ""
    ]
    if missing:
        error: str = f"Missing fields: {', '.join(missing)}"
        print(f"{error} for {username}")

        # Add failure info to report
        user_report.update({
            "Error": error,
            "Loan USD": None,
            "Down Payment USD": None,
            "Loan EUR": None,
            "Initial Deposit (Corrected)": None,
            "Initial Deposit Used": None,
        })
        report.append(user_report)
        continue

    # -------------------------
    # Step 2: Validate Initial Deposit
    # -------------------------
    correction_applied: bool = False
    try:
        initial_deposit: float = float(user["Initial Deposit"])
        if pd.isna(initial_deposit) or initial_deposit <= 0:
            corrected_deposit: float = 100.0
            correction_applied = True
        else:
            corrected_deposit = initial_deposit
    except Exception:
        corrected_deposit = 100.0
        correction_applied = True

    # Track corrected deposit separately
    user_report["Initial Deposit (Corrected)"] = corrected_deposit if correction_applied else None
    user_report["Initial Deposit Used"] = corrected_deposit

    # Down payment = 20% of corrected deposit
    down_payment: float = round(corrected_deposit * 0.2, 2)

    # =====================================================
    # Step 3: Start Browser Session
    # =====================================================
    options = webdriver.ChromeOptions()
    options.add_argument("--headless=new")  # Run without opening a window

    driver: webdriver.Chrome = webdriver.Chrome(
        service=Service(ChromeDriverManager().install()),
        options=options
    )
    wait: WebDriverWait = WebDriverWait(driver, 10)
    driver.get("https://parabank.parasoft.com/parabank/index.htm")

    try:
        # -------------------------
        # Step 4: Attempt Registration
        # -------------------------
        wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "Register"))).click()

        driver.find_element(By.ID, "customer.firstName").send_keys(user["First Name"])
        driver.find_element(By.ID, "customer.lastName").send_keys(user["Last Name"])
        driver.find_element(By.ID, "customer.address.street").send_keys(user["Address"])
        driver.find_element(By.ID, "customer.address.city").send_keys(user["City"])
        driver.find_element(By.ID, "customer.address.state").send_keys(user["State"])
        driver.find_element(By.ID, "customer.address.zipCode").send_keys(str(user["Zip Code"]))
        driver.find_element(By.ID, "customer.phoneNumber").send_keys(str(user["Phone Number"]))
        driver.find_element(By.ID, "customer.ssn").send_keys(str(user["SSN"]))
        driver.find_element(By.ID, "customer.username").send_keys(username)
        driver.find_element(By.ID, "customer.password").send_keys(password)
        driver.find_element(By.ID, "repeatedPassword").send_keys(password)
        driver.find_element(By.XPATH, "//input[@value='Register']").click()

        # If registration is successful
        if "Welcome" in driver.page_source:
            print(f"Registration successful for {username}")
            user_report["Registration Status"] = "Success"
            user_report["Login Status"] = "Success"
        else:
            # Fallback: try login
            print(f"Username {username} exists, trying login...")
            user_report["Registration Status"] = "Exists"

            driver.get("https://parabank.parasoft.com/parabank/index.htm")
            user_box = wait.until(EC.presence_of_element_located((By.NAME, "username")))
            user_box.clear()
            user_box.send_keys(username)

            pass_box = driver.find_element(By.NAME, "password")
            pass_box.clear()
            pass_box.send_keys(password)

            driver.find_element(By.XPATH, "//input[@value='Log In']").click()

            if "Welcome" in driver.page_source or "Log Out" in driver.page_source:
                print(f"Login successful for {username}")
                user_report["Login Status"] = "Success"
            else:
                error = "Login failed"
                print(f"{error} for {username}")
                user_report["Login Status"] = "Failed"
                user_report["Error"] = error
                report.append(user_report)
                continue

        # -------------------------
        # Step 5: Open New Account
        # -------------------------
        wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "Open New Account"))).click()
        wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@value='Open New Account']"))).click()
        user_report["Account Opened"] = "Success"
        print(f"New account opened for {username}")

        # -------------------------
        # Step 6: Request Loan
        # -------------------------
        wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "Request Loan"))).click()
        driver.find_element(By.ID, "amount").send_keys(str(LOAN_AMOUNT_USD))
        driver.find_element(By.ID, "downPayment").send_keys(str(down_payment))
        driver.find_element(By.XPATH, "//input[@value='Apply Now']").click()

        # Convert loan to EUR and update report
        loan_eur: float = round(LOAN_AMOUNT_USD * USD_TO_EUR, 2)
        user_report.update({
            "Loan USD": LOAN_AMOUNT_USD,
            "Down Payment USD": down_payment,
            "Loan EUR": loan_eur,
            "Loan Requested": "Success"
        })
        print(f"Loan requested for {username}: {LOAN_AMOUNT_USD} USD (~{loan_eur} EUR)")

        # -------------------------
        # Step 7: Logout
        # -------------------------
        wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "Log Out"))).click()

    except Exception as e:
        # Capture unexpected errors for this user
        error: str = f"Unexpected error: {e}"
        print(f"{error} for {username}")
        user_report.update({
            "Error": error,
            "Loan Requested": "Failed"
        })

    finally:
        # Always close browser and append user result
        report.append(user_report)
        driver.quit()

# =====================================================
# Save Final Report
# =====================================================
df_report: pd.DataFrame = pd.DataFrame(report)

# Reorder columns for clarity
cols_order: List[str] = [
    "Username", "First Name", "Last Name",
    "DOB", "Debit Card", "CVV",  # Include extra details
    "Registration Status", "Login Status", "Account Opened", "Loan Requested",
    "Initial Deposit", "Initial Deposit (Corrected)", "Initial Deposit Used",
    "Loan USD", "Down Payment USD", "Loan EUR",
    "Error"
]
df_report = df_report[[c for c in cols_order if c in df_report.columns]]

# Save to Excel
df_report.to_excel("Parabank_Report.xlsx", index=False)
print("Automation completed. Report saved as Parabank_Report.xlsx")
