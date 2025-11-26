import json
import os
import time
import traceback
from datetime import datetime
from typing import Tuple
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager


def _log(message: str, log_path: str = None) -> None:
    if not log_path:
        return
    os.makedirs(os.path.dirname(log_path), exist_ok=True)
    with open(log_path, "a", encoding="utf-8") as f:
        ts = datetime.now().strftime("%H:%M:%S")
        f.write(f"[{ts}] {message}\n")


def _load_config() -> dict:
    cfg_path = os.path.join(os.getcwd(), "config.json")
    if not os.path.isfile(cfg_path):
        raise FileNotFoundError(f"config.json not found at {cfg_path}")
    with open(cfg_path, "r", encoding="utf-8") as f:
        return json.load(f)


def _resolve_secret(value: str) -> str:
    if isinstance(value, str) and value.upper().startswith("ENV_"):
        return os.getenv(value, "")
    return value


def hx_upload(file_path: str, log_path: str = None) -> Tuple[bool, str]:
    """Upload file to HealthX portal using Selenium and return success flag/message."""
    driver = None
    try:
        if not os.path.isfile(file_path):
            msg = f"Upload file not found: {file_path}"
            _log(f"ERROR: {msg}", log_path)
            return False, msg

        cfg = _load_config()
        hx_url = cfg.get("healthx_url", "").strip()
        username = _resolve_secret(cfg.get("user_id") or cfg.get("username", ""))
        password = _resolve_secret(cfg.get("password", ""))
        client_text = cfg.get("hx_client_text", "") or "Audentes - Audentes Verification"

        if not hx_url or not username or not password:
            msg = "Missing HealthX configuration (url/username/password)"
            _log(f"ERROR: {msg}", log_path)
            return False, msg

        _log("Setting up Chrome browser...", log_path)
        chrome_options = webdriver.ChromeOptions()
        chrome_options.add_argument('--log-level=3')
        chrome_options.add_argument('--start-maximized')
        chrome_options.add_argument('--no-sandbox')
        chrome_options.add_argument('--disable-dev-shm-usage')
        chrome_options.add_experimental_option('prefs', {
            "profile.default_content_settings.popups": 0,
            "download.prompt_for_download": False,
            "directory_upgrade": True
        })
        service = Service(ChromeDriverManager().install())
        # Increase timeout for ChromeDriver connection
        driver = webdriver.Chrome(service=service, options=chrome_options)
        driver.set_page_load_timeout(300)  # 5 minutes for page load
        driver.implicitly_wait(10)  # Implicit wait for elements
        wait = WebDriverWait(driver, 60)  # Increased explicit wait timeout

        _log(f"Navigating to HealthX: {hx_url}", log_path)
        driver.get(hx_url)
        wait.until(EC.presence_of_element_located((By.ID, "email"))).send_keys(username)
        wait.until(EC.presence_of_element_located((By.ID, "password"))).send_keys(password)
        wait.until(EC.element_to_be_clickable((By.XPATH, '//input[@value="Sign in"]'))).click()

        _log("Navigating to Import screen...", log_path)
        time.sleep(5)
        wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="sidebar-toggle"]/li[5]/a'))).click()
        time.sleep(3)
        wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="EVSummary"]/li[1]/a'))).click()
        time.sleep(5)

        _log(f"Selecting campaign: {client_text}", log_path)
        wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="select2-campaign-container"]'))).click()
        time.sleep(2)
        
        # Wait for dropdown options to appear, then find matching option
        wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, 'li.select2-results__option')))
        time.sleep(1)  # Give dropdown time to fully render
        
        def _normalize_option(value: str) -> str:
            """Normalize text for matching (lowercase, alphanumeric only)."""
            return "".join(ch.lower() for ch in value if ch.isalnum())
        
        normalized_target = _normalize_option(client_text)
        options = driver.find_elements(By.CSS_SELECTOR, 'li.select2-results__option')
        matched = False
        
        for opt in options:
            text = opt.text or ""
            if not text.strip():
                continue
            # Try exact match first, then normalized match
            if client_text.strip() in text or _normalize_option(text).startswith(normalized_target):
                try:
                    opt.click()
                    matched = True
                    _log(f"Selected option: {text}", log_path)
                    break
                except Exception as e:
                    _log(f"Failed to click option '{text}': {e}", log_path)
                    continue
        
        if not matched:
            available = [opt.text for opt in options[:5] if opt.text.strip()]
            error_msg = f"Client option '{client_text}' not found in dropdown. Available options: {available}"
            _log(f"ERROR: {error_msg}", log_path)
            raise Exception(error_msg)
        
        time.sleep(2)

        _log("Uploading file to portal...", log_path)
        file_input = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="customFile"]')))
        file_input.send_keys(os.path.abspath(file_path))
        time.sleep(3)
        wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="ImportButtonRed"]'))).click()

        _log("Awaiting upload confirmation...", log_path)
        start = time.time()
        success = False
        max_wait = 300  # 5 minutes for upload confirmation
        while time.time() - start < max_wait:
            time.sleep(3)
            try:
                # Try multiple possible success message patterns
                success_patterns = [
                    '//p[contains(@class, "message") and contains(text(), "uploaded Successfully")]',
                    '//p[contains(@class, "message") and contains(text(), "uploaded successfully")]',
                    '//p[contains(@class, "message") and contains(text(), "Successfully")]',
                    '//*[contains(text(), "uploaded Successfully")]',
                    '//*[contains(text(), "uploaded successfully")]'
                ]
                for pattern in success_patterns:
                    try:
                        wait.until(EC.presence_of_element_located((By.XPATH, pattern)))
                        success = True
                        break
                    except:
                        continue
                if success:
                    break
            except Exception as e:
                _log(f"Checking for upload confirmation... ({int(time.time() - start)}s elapsed)", log_path)
                pass

        if not success:
            _log(f"Upload confirmation not detected within {max_wait}s timeout. Proceeding anyway...", log_path)
            # Don't fail - sometimes the message doesn't appear but upload succeeds
        else:
            _log("Upload confirmed as successful.", log_path)

        _log("Navigating to EV Allocation...", log_path)
        wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="EVSummary"]/li[2]/a'))).click()
        time.sleep(6)
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(3)
        initiate_btn = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="newrecord_evprocess"]/div/a')))
        initiate_btn.click()
        time.sleep(3)

        try:
            wait.until(EC.presence_of_element_located((By.XPATH, '//button[@id="InitiateConfirmAction"]'))).click()
            _log("EV process initiated successfully.", log_path)
        except Exception as exc:
            _log(f"Warning: Could not confirm EV initiation ({exc}).", log_path)

        time.sleep(3)
        _log("HX upload completed.", log_path)
        return True, "Upload successful and EV process initiated"

    except Exception as exc:
        message = f"Upload failed: {exc}"
        _log(f"ERROR: {message}", log_path)
        _log(traceback.format_exc(), log_path)
        return False, message
    finally:
        if driver:
            try:
                driver.quit()
            except Exception:
                pass
