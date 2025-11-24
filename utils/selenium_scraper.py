"""
Selenium-based web scraper for JavaScript-heavy websites
Falls back from BeautifulSoup/requests when needed
"""

import logging
import time
import json
import re
from typing import Dict, List, Optional, Callable, TYPE_CHECKING
from urllib.parse import urljoin, urlparse
from bs4 import BeautifulSoup

logger = logging.getLogger(__name__)

# Import By for type hints only
if TYPE_CHECKING:
    from selenium.webdriver.common.by import By

try:
    from selenium import webdriver
    from selenium.webdriver.chrome.service import Service
    from selenium.webdriver.chrome.options import Options
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    from selenium.common.exceptions import TimeoutException, WebDriverException
    SELENIUM_AVAILABLE = True
except ImportError:
    SELENIUM_AVAILABLE = False
    # Create a dummy By class for type hints when Selenium is not available
    class By:
        CSS_SELECTOR = "css selector"
        XPATH = "xpath"
        TAG_NAME = "tag name"
        ID = "id"
        CLASS_NAME = "class name"
    logger.warning("Selenium not installed. Install with: pip install selenium webdriver-manager")

try:
    from webdriver_manager.chrome import ChromeDriverManager
    WEBDRIVER_MANAGER_AVAILABLE = True
except ImportError:
    WEBDRIVER_MANAGER_AVAILABLE = False
    logger.warning("webdriver-manager not installed. Install with: pip install webdriver-manager")


class SeleniumScraper:
    """Selenium-based scraper for JavaScript-heavy websites"""
    
    def __init__(self, headless: bool = True, timeout: int = 30):
        """
        Initialize Selenium scraper
        
        Args:
            headless: Run browser in headless mode
            timeout: Page load timeout in seconds
        """
        self.headless = headless
        self.timeout = timeout
        self.driver = None
        self._init_driver()
    
    def _init_driver(self):
        """Initialize Chrome WebDriver"""
        if not SELENIUM_AVAILABLE:
            raise ImportError("Selenium is not installed. Install with: pip install selenium")
        
        try:
            chrome_options = Options()
            if self.headless:
                chrome_options.add_argument('--headless')
            chrome_options.add_argument('--no-sandbox')
            chrome_options.add_argument('--disable-dev-shm-usage')
            chrome_options.add_argument('--disable-blink-features=AutomationControlled')
            chrome_options.add_argument('--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36')
            chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
            chrome_options.add_experimental_option('useAutomationExtension', False)
            
            if WEBDRIVER_MANAGER_AVAILABLE:
                try:
                    service = Service(ChromeDriverManager().install())
                except Exception as e:
                    logger.warning(f"ChromeDriverManager failed: {e}, trying without version detection")
                    # Try to use ChromeDriverManager without version detection
                    try:
                        from webdriver_manager.chrome import ChromeDriverManager
                        from webdriver_manager.core.os_manager import ChromeType
                        manager = ChromeDriverManager()
                        # Force a specific version or skip version detection
                        service = Service(manager.install())
                    except Exception as e2:
                        logger.error(f"Failed to initialize ChromeDriverManager: {e2}")
                        raise ImportError(f"Cannot initialize Chrome WebDriver: {e2}")
            else:
                service = Service()  # Use system ChromeDriver
            
            self.driver = webdriver.Chrome(service=service, options=chrome_options)
            self.driver.implicitly_wait(10)
            self.driver.set_page_load_timeout(self.timeout)
            
            logger.info("Selenium WebDriver initialized successfully")
        except Exception as e:
            logger.error(f"Failed to initialize Selenium WebDriver: {e}")
            raise
    
    def get_page(self, url: str, wait_for_selector: Optional[str] = None, wait_time: int = 10) -> BeautifulSoup:
        """
        Load a page and return BeautifulSoup object
        
        Args:
            url: URL to load
            wait_for_selector: Optional CSS selector to wait for before returning
            wait_time: Maximum wait time in seconds
            
        Returns:
            BeautifulSoup object of the page
        """
        if not self.driver:
            self._init_driver()
        
        try:
            logger.info(f"Loading page: {url}")
            self.driver.get(url)
            
            # Wait for optional selector
            if wait_for_selector:
                try:
                    WebDriverWait(self.driver, wait_time).until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, wait_for_selector))
                    )
                except TimeoutException:
                    logger.warning(f"Selector {wait_for_selector} not found, continuing anyway")
            
            # Wait for page to be ready
            time.sleep(2)  # Additional wait for JavaScript
            
            # Get page source and parse
            page_source = self.driver.page_source
            soup = BeautifulSoup(page_source, 'html.parser')
            
            return soup
            
        except TimeoutException:
            logger.error(f"Timeout loading page: {url}")
            raise
        except Exception as e:
            logger.error(f"Error loading page {url}: {e}")
            raise
    
    def find_elements(self, by: By, value: str) -> List:
        """Find elements using Selenium"""
        if not self.driver:
            self._init_driver()
        
        try:
            return self.driver.find_elements(by, value)
        except Exception as e:
            logger.error(f"Error finding elements: {e}")
            return []
    
    def click_element(self, by: By, value: str, wait_time: int = 10):
        """Click an element"""
        if not self.driver:
            self._init_driver()
        
        try:
            element = WebDriverWait(self.driver, wait_time).until(
                EC.element_to_be_clickable((by, value))
            )
            element.click()
            time.sleep(1)  # Wait for action to complete
        except TimeoutException:
            logger.error(f"Timeout waiting for element to be clickable: {value}")
            raise
    
    def scroll_to_bottom(self, pause_time: float = 2.0):
        """Scroll to bottom of page to load dynamic content"""
        if not self.driver:
            self._init_driver()
        
        try:
            last_height = self.driver.execute_script("return document.body.scrollHeight")
            
            while True:
                # Scroll down
                self.driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
                time.sleep(pause_time)
                
                # Calculate new scroll height
                new_height = self.driver.execute_script("return document.body.scrollHeight")
                
                if new_height == last_height:
                    break
                last_height = new_height
            
            logger.info("Scrolled to bottom of page")
        except Exception as e:
            logger.error(f"Error scrolling page: {e}")
    
    def execute_script(self, script: str):
        """Execute JavaScript on the page"""
        if not self.driver:
            self._init_driver()
        
        try:
            return self.driver.execute_script(script)
        except Exception as e:
            logger.error(f"Error executing script: {e}")
            return None
    
    def wait_for_element(self, by: By, value: str, timeout: int = 10):
        """Wait for an element to appear"""
        if not self.driver:
            self._init_driver()
        
        try:
            return WebDriverWait(self.driver, timeout).until(
                EC.presence_of_element_located((by, value))
            )
        except TimeoutException:
            logger.warning(f"Timeout waiting for element: {value}")
            return None
    
    def close(self):
        """Close the browser"""
        if self.driver:
            try:
                self.driver.quit()
                self.driver = None
                logger.info("Selenium WebDriver closed")
            except Exception as e:
                logger.error(f"Error closing WebDriver: {e}")
    
    def __enter__(self):
        """Context manager entry"""
        return self
    
    def __exit__(self, exc_type, exc_val, exc_tb):
        """Context manager exit"""
        self.close()
    
    def __del__(self):
        """Cleanup on deletion"""
        self.close()


def scrape_with_fallback(url: str, 
                         requests_scraper_func: Callable,
                         fallback_to_selenium: bool = True,
                         selenium_wait_selector: Optional[str] = None) -> Dict:
    """
    Try scraping with requests/BeautifulSoup first, fall back to Selenium if needed
    
    Args:
        url: URL to scrape
        requests_scraper_func: Function that takes URL and returns scraped data using requests
        fallback_to_selenium: Whether to use Selenium if requests fails
        selenium_wait_selector: CSS selector to wait for when using Selenium
        
    Returns:
        Dictionary with scraped data or error
    """
    # Try requests first
    try:
        logger.info(f"Attempting to scrape {url} with requests/BeautifulSoup")
        result = requests_scraper_func(url)
        if result and 'error' not in result:
            logger.info(f"Successfully scraped {url} with requests/BeautifulSoup")
            return result
        else:
            logger.info(f"Requests scraping failed or returned error, trying Selenium")
    except Exception as e:
        logger.warning(f"Requests scraping failed: {e}, trying Selenium")
    
    # Fall back to Selenium if available
    if fallback_to_selenium and SELENIUM_AVAILABLE:
        try:
            logger.info(f"Attempting to scrape {url} with Selenium")
            scraper = SeleniumScraper(headless=True)
            
            try:
                soup = scraper.get_page(url, wait_for_selector=selenium_wait_selector)
                
                # Try to extract data using similar logic
                # This is a basic implementation - should be customized per site
                result = {
                    'success': True,
                    'soup': soup,
                    'url': url,
                    'method': 'selenium'
                }
                
                return result
            finally:
                scraper.close()
                
        except Exception as e:
            logger.error(f"Selenium scraping also failed: {e}")
            return {'error': f'Both requests and Selenium scraping failed: {str(e)}'}
    else:
        return {'error': 'Requests scraping failed and Selenium not available'}





