import json
import requests
from datetime import datetime, timedelta

def fetch_match_data(player1, player2, num_pages=10):
    base_url = f"https://football.esportsbattle.com/api/participants/{player1}/compare/{player2}/matches?page={{}}"
    data = []
    for page in range(1, num_pages + 1):
        url = base_url.format(page)
        response = requests.get(url)
        if response.status_code == 200:
            json_data = response.json()
            matches = json_data.get("matches", [])
            for match in matches:
                data.append({
                    "Match ID": match["id"],
                    "Date": match["date"],
                    "Player 1": match["participant1"]["nickname"],
                    "Team 1": match["participant1"]["team"]["token_international"],
                    "Score player 1": match["participant1"]["score"],
                    "Score player 1 1st half": match["participant1"].get("prevPeriodsScores", [0])[0],
                    "Player 2": match["participant2"]["nickname"],
                    "Team 2": match["participant2"]["team"]["token_international"],
                    "Score player 2 1st half": match["participant2"].get("prevPeriodsScores", [0])[0],
                    "Score player 2": match["participant2"]["score"],
                })
        else:
            print(f"Failed to fetch page {page}: {response.status_code}")
    return data

def filter_games(data, player1, player2, days_back=None):
    filtered = []
    now = datetime.now()
    processed_game_ids = set()
    for game in data:
        try:
            game_id = game["Match ID"]
            if game_id in processed_game_ids:
                continue
            home_player = game["Player 1"]
            away_player = game["Player 2"]
            if ((home_player == player1 and away_player == player2) or
                (home_player == player2 and away_player == player1)):
                if days_back:
                    game_time = datetime.strptime(game["Date"], "%Y-%m-%dT%H:%M:%SZ")
                    if now - game_time <= timedelta(days=days_back):
                        filtered.append(game)
                        processed_game_ids.add(game_id)
                else:
                    filtered.append(game)
                    processed_game_ids.add(game_id)
        except (KeyError, IndexError):
            continue
    return filtered

def calculate_stats(games, player1):
    wins, draws, losses = 0, 0, 0
    total_games = len(games)
    if total_games == 0:
        return {"win": 0, "draw": 0, "loss": 0, "total": 0}
    for game in games:
        home_player = game["Player 1"]
        away_player = game["Player 2"]
        home_score = int(game["Score player 1"])
        away_score = int(game["Score player 2"])
        if home_player == player1:
            if home_score > away_score:
                wins += 1
            elif home_score == away_score:
                draws += 1
            else:
                losses += 1
        elif away_player == player1:
            if away_score > home_score:
                wins += 1
            elif home_score == away_score:
                draws += 1
            else:
                losses += 1
    win_percentage = round((wins / total_games) * 100, 2)
    draw_percentage = round((draws / total_games) * 100, 2)
    loss_percentage = round((losses / total_games) * 100, 2)
    return {
        "win": win_percentage,
        "draw": draw_percentage,
        "loss": loss_percentage,
        "total": total_games
    }

def calculate_average_goals_per_half_and_total(games, player1, player2):
    goals = {
        player1: {"first_half": 0, "second_half": 0},
        player2: {"first_half": 0, "second_half": 0},
        "total_goals": {"first_half": 0, "second_half": 0, "full_time": 0},
    }
    valid_games = 0  
    for game in games:
        try:
            home_player = game["Player 1"]
            away_player = game["Player 2"]
            if game["Score player 1"] is None or game["Score player 2"] is None:
                continue
            first_half_home = int(game["Score player 1 1st half"])
            first_half_away = int(game["Score player 2 1st half"])
            second_half_home = int(game["Score player 1"]) - first_half_home
            second_half_away = int(game["Score player 2"]) - first_half_away
            total_first_half = first_half_home + first_half_away
            total_second_half = second_half_home + second_half_away
            total_full_time = total_first_half + total_second_half
            goals["total_goals"]["first_half"] += total_first_half
            goals["total_goals"]["second_half"] += total_second_half
            goals["total_goals"]["full_time"] += total_full_time
            if home_player == player1:
                goals[player1]["first_half"] += first_half_home
                goals[player1]["second_half"] += second_half_home
                goals[player2]["first_half"] += first_half_away
                goals[player2]["second_half"] += second_half_away
            elif away_player == player1:
                goals[player1]["first_half"] += first_half_away
                goals[player1]["second_half"] += second_half_away
                goals[player2]["first_half"] += first_half_home
                goals[player2]["second_half"] += second_half_home
            valid_games += 1 
        except (KeyError, IndexError):
            continue
    if valid_games == 0:
        return {
            player1: {"first_half": 0, "second_half": 0},
            player2: {"first_half": 0, "second_half": 0},
            "total_goals": {"first_half": 0, "second_half": 0, "full_time": 0},
        }
    return {
        player1: {
            "first_half": round(goals[player1]["first_half"] / valid_games, 2),
            "second_half": round(goals[player1]["second_half"] / valid_games, 2),
        },
        player2: {
            "first_half": round(goals[player2]["first_half"] / valid_games, 2),
            "second_half": round(goals[player2]["second_half"] / valid_games, 2),
        },
        "total_goals": {
            "first_half": round(goals["total_goals"]["first_half"] / valid_games, 2),
            "second_half": round(goals["total_goals"]["second_half"] / valid_games, 2),
            "full_time": round(goals["total_goals"]["full_time"] / valid_games, 2),
        },
    }

def calculate_goal_thresholds(games):
    thresholds = [2.5, 2.75, 3.0, 3.25, 3.5, 3.75, 4.0, 4.25, 4.5, 4.75, 5.0, 5.25, 5.5, 5.75, 6.0, 6.25, 6.5, 6.75, 7.0, 7.25, 7.5, 7.75, 8.0, 8.25, 8.5]
    above_counts = {str(threshold): 0 for threshold in thresholds}
    below_counts = {str(threshold): 0 for threshold in thresholds}
    valid_games = 0  
    for game in games:
        try:
            if game["Score player 1"] is None or game["Score player 2"] is None:
                continue
            home_score = int(game["Score player 1"])
            away_score = int(game["Score player 2"])
            total_goals = home_score + away_score
            for threshold in thresholds:
                if total_goals > threshold:
                    above_counts[str(threshold)] += 1
                else:
                    below_counts[str(threshold)] += 1
            valid_games += 1  
        except (KeyError, IndexError):
            continue
    if valid_games == 0:
        return {
            "above": {str(threshold): 0 for threshold in thresholds},
            "below": {str(threshold): 0 for threshold in thresholds},
        }
    return {
        "above": {threshold: round((count / valid_games) * 100, 2) for threshold, count in above_counts.items()},
        "below": {threshold: round((count / valid_games) * 100, 2) for threshold, count in below_counts.items()},
    }

def save_stats_to_json(player1, player2, stats, all_games, output_file="data/player_stats_output.json"):
    output_data = {
        "player1": player1,
        "player2": player2,
        "stats": stats,
        "games": all_games,
    }
    with open(output_file, "w", encoding="utf-8") as file:
        json.dump(output_data, file, indent=4)
    print(f"Stats and games saved to {output_file}")

def filter_valid_games(games):
    valid_games = []
    for game in games:
        if game["Score player 1"] is not None and game["Score player 2"] is not None:
            valid_games.append(game)
    return valid_games

def extract_player_name(team_name):
    start = team_name.find("(")
    end = team_name.find(")")
    if start != -1 and end != -1:
        return team_name[start + 1:end]
    return None

def main():
    with open("data/tippmixpro_upcoming_games.json", "r") as file:
        player_pairs = json.load(file)
    all_results = []
    for game in player_pairs["games"]:
        home_team = game["home"]
        away_team = game["away"]
        player1 = extract_player_name(home_team)
        player2 = extract_player_name(away_team)
        if not player1 or not player2:
            print(f"Skipping invalid player names: {home_team} vs {away_team}")
            continue
        print(f"Processing {player1} vs {player2}...")
        data = fetch_match_data(player1, player2)
        filtered_games = filter_games(data, player1, player2)
        valid_games = filter_valid_games(filtered_games)
        valid_games_sorted = sorted(valid_games, key=lambda x: datetime.strptime(x["Date"], "%Y-%m-%dT%H:%M:%SZ"), reverse=True)
        games_25 = valid_games_sorted[:25]
        games_50 = valid_games_sorted[:50]
        games_30_days = filter_games(valid_games, player1, player2, days_back=30)
        stats_25 = calculate_stats(games_25, player1)
        goals_25 = calculate_average_goals_per_half_and_total(games_25, player1, player2)
        thresholds_25 = calculate_goal_thresholds(games_25)
        stats_50 = calculate_stats(games_50, player1)
        goals_50 = calculate_average_goals_per_half_and_total(games_50, player1, player2)
        thresholds_50 = calculate_goal_thresholds(games_50)
        stats_30_days = calculate_stats(games_30_days, player1)
        goals_30_days = calculate_average_goals_per_half_and_total(games_30_days, player1, player2)
        thresholds_30_days = calculate_goal_thresholds(games_30_days)
        combined_stats = {
            "past_25": {"win_draw_loss": stats_25, "avg_goals": goals_25, "goal_thresholds": thresholds_25},
            "past_50": {"win_draw_loss": stats_50, "avg_goals": goals_50, "goal_thresholds": thresholds_50},
            "past_30_days": {"win_draw_loss": stats_30_days, "avg_goals": goals_30_days, "goal_thresholds": thresholds_30_days},
        }
        all_games = []
        seen_game_ids = set()
        for game in games_25 + games_50 + games_30_days:
            game_id = game["Match ID"]
            if game_id not in seen_game_ids:
                all_games.append(game)
                seen_game_ids.add(game_id)
        output_file = f"data/{player1}_vs_{player2}_stats.json"
        save_stats_to_json(player1, player2, combined_stats, all_games, output_file)
        all_results.append({
            "home": home_team,
            "away": away_team,
            "stats": combined_stats,
        })
    with open("data/all_results.json", "w", encoding="utf-8") as file:
        json.dump(all_results, file, indent=4)
    print("All results saved to data/all_results.json")

if __name__ == "__main__":
    main()
