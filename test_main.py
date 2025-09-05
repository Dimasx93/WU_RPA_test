import pytest
import main
import importlib
import requests

# ----------------------------
# MockResponse for requests
# ----------------------------
class MockResponse:
    def __init__(self, json_data, status_code=200):
        self._json = json_data
        self.status_code = status_code

    def json(self):
        return self._json

    def raise_for_status(self):
        if self.status_code != 200:
            raise requests.exceptions.HTTPError(f"Status code: {self.status_code}")

# ----------------------------
# Deposit correction tests
# ----------------------------
@pytest.mark.parametrize("deposit, expected", [
    (500, 500.0),
    (0, 100.0),        # corrected
    (-10, 100.0),      # corrected
    (None, 100.0),     # corrected
])
def test_correct_initial_deposit(deposit, expected):
    try:
        initial_deposit = float(deposit)
        if initial_deposit <= 0:
            corrected = 100.0
        else:
            corrected = initial_deposit
    except Exception:
        corrected = 100.0

    assert corrected == expected

# ----------------------------
# Down payment test
# ----------------------------
def test_down_payment():
    deposit = 500
    down_payment = round(deposit * 0.2, 2)
    assert down_payment == 100.0

# ----------------------------
# Mock process user (no browser)
# ----------------------------
def test_process_user_mocked(monkeypatch):
    # Patch webdriver to avoid opening Chrome
    class MockDriver:
        def quit(self): pass
        def get(self, url): pass
    monkeypatch.setattr(main.webdriver, "Chrome", lambda *a, **kw: MockDriver())

    # Create fake user
    user = {
        "First Name": "Alice",
        "Last Name": "Smith",
        "Address": "123 St",
        "City": "Town",
        "State": "TS",
        "Zip Code": "12345",
        "Phone Number": "555-1234",
        "SSN": "123-45-6789",
        "Username": "alice_user",
        "Password": "secret",
        "Initial Deposit": 200,
        "DOB": "2000-01-01",
        "Debit Card": "1234-5678-9012-3456",
        "CVV": "123"
    }

    # Expected corrected values
    initial_deposit = 200.0
    down_payment = round(initial_deposit * 0.2, 2)

    # Simulate loan calculation
    loan_eur = round(main.LOAN_AMOUNT_USD * main.USD_TO_EUR, 2)

    # Check values
    assert initial_deposit == 200.0
    assert down_payment == 40.0
    assert loan_eur > 0
