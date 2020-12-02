from LeagueScouterPackage.LeagueScouter import LeagueScouter
import json

credentials = open('./credentials.json',)
credentials_data = json.load(credentials)

scouter = LeagueScouter("Worlds 2020 Main Event",credentials_data['username'],credentials_data['password'],"./chromedriver")

scouter.update_games_data()
scouter.get_competitive_stats()
scouter.draft_picks_and_bans()