# utils/selenium_helpers.py
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC

class text_to_be_present_in_element_either:
    """
    Custom expected condition for Selenium that waits until
    any of the specified texts are present in the element.
    """
    def __init__(self, locator, texts):
        self.locator = locator
        self.texts = texts

    def __call__(self, driver):
        try:
            element = driver.find_element(*self.locator)
            element_text = element.text
            for text in self.texts:
                if text in element_text:
                    return True
            return False
        except Exception:
            return False

def select_ticket_card(wait, ticket_type):
    """
    Selects the ticket card based on the ticket_type string.

    Args:
        wait: Selenium WebDriverWait instance.
        ticket_type: The ticket type to select (e.g., "Desktops / Laptops", "Projectors / Speakers", "Chromebooks").
    """
    ticket_type_to_xpath = {
        "Desktops / Laptops": "//p[text()='Desktop / Laptops']/ancestor::div[contains(@class, 'gh-flat-corner')]",
        "Projectors / Speakers": "//p[text()='Projectors / Speakers']/ancestor::div[contains(@class, 'gh-flat-corner') and contains(@class, 'card')]",
        # "Chromebooks": "//p[text()='Chromebooks']/ancestor::div[contains(@class, 'gh-flat-corner')]" # Untested
    }

    if ticket_type not in ticket_type_to_xpath:
        raise ValueError(f"Unsupported ticket type: {ticket_type}")

    xpath_selector = ticket_type_to_xpath[ticket_type]
    card_element = wait.until(
        EC.element_to_be_clickable((By.XPATH, xpath_selector))
    )
    card_element.click()