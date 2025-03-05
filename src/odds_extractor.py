import json
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from bs4 import BeautifulSoup

def extract_market_titles_and_odds(url):
    options = webdriver.ChromeOptions()
    options.add_argument('--headless')
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-dev-shm-usage')

    options = webdriver.ChromeOptions()
    options.binary_location = "/usr/bin/chromium"  # Path to Chromium
    service = Service("/usr/bin/chromedriver")  # Path to ChromeDriver

# Create the WebDriver
    driver = webdriver.Chrome(service=service, options=options)

    try:
        driver.get(url)
        WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.ID, "SportsIframe")))
        iframe = driver.find_element(By.ID, "SportsIframe")
        driver.switch_to.frame(iframe)
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CLASS_NAME, "OddsButton")))
        iframe_content = driver.page_source
        soup = BeautifulSoup(iframe_content, 'html.parser')
        market_data = []
        for market in soup.find_all('article', class_='Market'):
            market_classes = market.get('class', [])
            market_id = None
            for cls in market_classes:
                if cls.startswith('Market--Id-'):
                    market_id = cls.split('Market--Id-')[-1]
                    break
            market_part = None
            for cls in market_classes:
                if cls.startswith('Market--Part-'):
                    market_part = cls.split('Market--Part-')[-1]
                    break
            market_title = market.find('span', class_='Market__CollapseText')
            market_title = market_title.get_text(strip=True) if market_title else "Unknown Market"
            if market_part == "2255":
                market_title += " - Full Game"
            elif market_part == "2256":
                market_title += " - First Half"
            market_odds = []
            if market_id in ["69", "9", "45", "11"]:
                for button in market.find_all('button', class_='OddsButton'):
                    text = button.find('span', class_='OddsButton__Text')
                    odds = button.find('span', class_='OddsButton__Odds')
                    if text and odds:
                        team = text.get_text(strip=True)
                        odds_value = odds.get_text(strip=True)
                        market_odds.append({
                            "team": team,
                            "odds": odds_value
                        })
            elif market_id in ["47", "77"]:
                for group in market.find_all('ul', class_='Market__OddsGroup'):
                    title = group.find('li', class_='Market__OddsGroupTitle')
                    if title:
                        line = title.get_text(strip=True)
                        odds_items = group.find_all('li', class_='Market__OddsGroupItem')
                        if len(odds_items) >= 2:
                            market_odds.append({
                                "line": line,
                                "over": odds_items[0].get_text(strip=True),
                                "under": odds_items[1].get_text(strip=True)
                            })
            market_data.append({
                "market_id": market_id,
                "market_part": market_part,
                "market_title": market_title,
                "odds": market_odds
            })
        return market_data
    finally:
        driver.quit()

def extract_odds():
    with open('data/tippmixpro_upcoming_games.json', 'r') as file:
        data = json.load(file)
    games_with_market_data = []
    for game in data['games']:
        print(f"Extracting market data for {game['home']} vs {game['away']}...")
        market_data = extract_market_titles_and_odds(game['link'])
        game_data = {
            'home': game['home'],
            'away': game['away'],
            'link': game['link'],
            'market_data': market_data
        }
        games_with_market_data.append(game_data)
    with open('data/games_with_odds.json', 'w') as outfile:
        json.dump(games_with_market_data, outfile, indent=4, ensure_ascii=False)
    print("✅ Market data extraction complete. Check 'data/games_with_odds.json' for the results.")

if __name__ == "__main__":
    extract_odds()
