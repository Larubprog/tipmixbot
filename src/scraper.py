import json
import asyncio
from playwright.async_api import async_playwright
from bs4 import BeautifulSoup

WATCHED_LEAGUES = [
    'esports battle',
    'escocer battle',
    'esport pro club',
    'cyber live! arena (2x5 min)',  
]
OUTPUT_FILE = 'data/tippmixpro_upcoming_games.json'
URL = "https://sports2.tippmixpro.hu/hu/fogadas/e-labdarugas/121/osszes/0/helyszin"

async def scroll_to_bottom(page):
    last_height = await page.evaluate('document.body.scrollHeight')
    while True:
        await page.evaluate('window.scrollTo(0, document.body.scrollHeight)')
        await asyncio.sleep(2)  
        new_height = await page.evaluate('document.body.scrollHeight')
        if new_height == last_height:
            break
        last_height = new_height

async def scrape_tippmix():
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        page = await browser.new_page()

        await page.goto(URL, timeout=60000)
        await page.wait_for_selector('.MatchList__Group')
        
        await scroll_to_bottom(page)
        
        content = await page.content()
        soup = BeautifulSoup(content, "lxml")

        leagues = soup.select('.MatchList__Group')
        data = {"games": []}

        for league in leagues:
            league_name = league.select_one('.MatchListGroup__Tournament').text.strip().lower()
            if not any(watched in league_name for watched in WATCHED_LEAGUES):
                continue

            events = league.select('.EventItem')
            for event in events:
                if 'EventItem--isLive' in event.get('class', []):
                    continue
                
                match_url = event.select_one('a')['href'] + "/all"
                home_team = event.select_one('.Details__Participant--Home .Details__ParticipantName').text.strip()
                away_team = event.select_one('.Details__Participant--Away .Details__ParticipantName').text.strip()
                
                print(f"Extracting match: {home_team} vs {away_team}")  
                
                event_info = {
                    "home": home_team,
                    "away": away_team,
                    "link": match_url.replace('/hu/', 'https://www.tippmixpro.hu/hu/fogadas/i/')
                }
                data["games"].append(event_info)

        with open(OUTPUT_FILE, 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=2, ensure_ascii=False)

        print(f"✅ Data successfully saved to {OUTPUT_FILE}")
        await browser.close()

if __name__ == "__main__":
    asyncio.run(scrape_tippmix())