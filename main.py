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
# Load Users
# =====================================================
users: pd.DataFrame = pd.read_csv("ParaBank users.csv")

# Required fields for registration
REQUIRED_FIELDS: List[str] = [
    "First Name", "Last Name", "Address", "City", "State", "Zip Code",
    "Phone Number", "SSN", "Username", "Password"
]

# Retain necessary columns only
users = users[
    [
        "First Name", "Last Name", "Address", "City", "State", "Zip Code",
        "Phone Number", "SSN", "Username", "Password",
        "Initial Deposit", "DOB", "Debit Card", "CVV"
    ]
]

# =====================================================
# Get Live USD → EUR rate
# =====================================================
try:
    resp = requests.get("https://open.er-api.com/v6/latest/USD", timeout=10)
    resp.raise_for_status()
    rate = resp.json().get("rates", {}).get("EUR")
    if rate is None:
        raise ValueError("EUR rate not found in API response")
    USD_TO_EUR: float = rate
    print(f"Live USD→EUR rate: 1 USD = {USD_TO_EUR:.2f} EUR")
except Exception as ex:
    print(f"⚠ Failed to fetch live rate, using fallback 0.92. Error: {ex}")
    USD_TO_EUR: float = 0.86

# =====================================================
# Loan Setup
# =====================================================
LOAN_AMOUNT_USD: int = 10000
report: List[Dict[str, Any]] = []

# =====================================================
# Process Each Customer
# =====================================================
for _, user in users.iterrows():
    username: str = str(user["Username"])
    password: str = str(user["Password"])

    # Keep all CSV fields
    user_report: Dict[str, Any] = user.to_dict()

    # -------------------------
    # Validate required fields
    # -------------------------
    missing: List[str] = [
        f for f in REQUIRED_FIELDS if pd.isna(user[f]) or str(user[f]).strip() == ""
    ]
    if missing:
        error = f"Missing fields: {', '.join(missing)}"
        print(f"{error} for {username}")

        user_report.update({
            "Error": error,
            "Loan USD": None,
            "Down Payment USD": None,
            "Loan EUR": None,
            "Initial Deposit (Corrected)": None
        })
        report.append(user_report)
        continue

    # -------------------------
    # Ensure Initial Deposit is valid
    # Default to 100 if null/0
    # -------------------------
    try:
        initial_deposit = float(user["Initial Deposit"])
        if pd.isna(initial_deposit) or initial_deposit <= 0:
            corrected_deposit = 100.0
        else:
            corrected_deposit = initial_deposit
    except Exception:
        corrected_deposit = 100.0

    # Save corrected deposit in report
    user_report["Initial Deposit (Corrected)"] = corrected_deposit

    # Down payment = 20% of corrected deposit
    down_payment: float = round(corrected_deposit * 0.2, 2)

    # =====================================================
    # Start Browser for this customer
    # =====================================================
    options = webdriver.ChromeOptions()
    options.add_argument("--headless=new")  # Run in headless mode

    driver = webdriver.Chrome(
        service=Service(ChromeDriverManager().install()),
        options=options
    )
    wait = WebDriverWait(driver, 10)
    driver.get("https://parabank.parasoft.com/parabank/index.htm")

    try:
        # -------------------------
        # Registration Attempt
        # -------------------------
        wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "Register"))).click()

        # Fill registration form
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

        # -------------------------
        # If registration failed → login instead
        # -------------------------
        if "Welcome" not in driver.page_source:
            print(f"Username {username} exists. Trying login...")
            driver.get("https://parabank.parasoft.com/parabank/index.htm")

            # Login
            user_box = wait.until(EC.presence_of_element_located((By.NAME, "username")))
            user_box.clear()
            user_box.send_keys(username)

            pass_box = driver.find_element(By.NAME, "password")
            pass_box.clear()
            pass_box.send_keys(password)

            driver.find_element(By.XPATH, "//input[@value='Log In']").click()

            try:
                wait.until(
                    EC.any_of(
                        EC.presence_of_element_located((By.LINK_TEXT, "Log Out")),
                        EC.text_to_be_present_in_element((By.TAG_NAME, "body"), "Welcome")
                    )
                )
                print(f"Logged in as {username}")
            except Exception:
                error = "Login failed"
                print(f"{error} for {username}")

                user_report.update({
                    "Error": error,
                    "Loan USD": LOAN_AMOUNT_USD,
                    "Down Payment USD": down_payment,
                    "Loan EUR": None
                })
                report.append(user_report)
                continue

        # -------------------------
        # Open New Account
        # -------------------------
        wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "Open New Account"))).click()
        wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@value='Open New Account']"))).click()

        # -------------------------
        # Request Loan
        # -------------------------
        wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "Request Loan"))).click()
        driver.find_element(By.ID, "amount").send_keys(str(LOAN_AMOUNT_USD))
        driver.find_element(By.ID, "downPayment").send_keys(str(down_payment))
        driver.find_element(By.XPATH, "//input[@value='Apply Now']").click()

        # Loan calculation
        loan_eur: float = round(LOAN_AMOUNT_USD * USD_TO_EUR, 2)

        user_report.update({
            "Error": None,
            "Loan USD": LOAN_AMOUNT_USD,
            "Down Payment USD": down_payment,
            "Loan EUR": loan_eur
        })
        report.append(user_report)

        print(f"Loan for {username}: {LOAN_AMOUNT_USD} USD (~{loan_eur} EUR), down payment {down_payment} USD")

        # Logout
        wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "Log Out"))).click()

    except Exception as e:
        error = f"Unexpected error: {e}"
        print(f"{error} for {username}")

        user_report.update({
            "Error": error,
            "Loan USD": LOAN_AMOUNT_USD,
            "Down Payment USD": down_payment,
            "Loan EUR": None
        })
        report.append(user_report)

    finally:
        driver.quit()

# =====================================================
# Save Final Report
# =====================================================
df_report = pd.DataFrame(report)
df_report.to_excel("Parabank_Report.xlsx", index=False)

print("Automation completed. Report saved as Parabank_Report.xlsx")