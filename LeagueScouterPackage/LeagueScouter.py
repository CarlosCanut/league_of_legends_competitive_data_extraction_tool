import mwclient
import pprint
import json
import os
import time
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import cv2
import seaborn as sns
import matplotlib.pyplot as plt
import numpy as np
import math
import xlsxwriter
import requests


##########################################
###########   Class Scouter   ############
##########################################
class LeagueScouter:
    tournament_name = ""
    username = ""
    password = ""
    chromedriver_path = ""
    # games data paths
    game_data_path = ""
    post_game_path = ""
    timeline_path = ""
    picks_and_bans_path = ""
    # analysis paths
    analysis_path = ""
    draft_stats_path = ""
    draft_picks_and_bans_path = ""
    raw_stats_path = ""
    heatmaps_path = ""
    histogram_path = ""


    #################################################################
    # creates all the paths and sets the configuration
    #################################################################
    def __init__(self, tournament_name, username, password, chromedriver_path):
        self.tournament_name = tournament_name
        self.username = username
        self.password = password
        self.chromedriver_path = chromedriver_path
        # game data paths
        self.game_data_path = "./Game_Data"
        self.post_game_path = self.game_data_path + "/Post_Game"
        self.timeline_path = self.game_data_path + "/Timeline"
        self.picks_and_bans_path = self.game_data_path + "/Picks_and_Bans"
        # creates game data paths if needed
        old_game_data_path = os.path.isdir(self.game_data_path)
        old_post_game_path = os.path.isdir(self.post_game_path)
        old_timeline_path = os.path.isdir(self.timeline_path)
        old_picks_and_bans_path = os.path.isdir(self.picks_and_bans_path)
        if not old_game_data_path:
            os.makedirs(self.game_data_path)
        else:
            pass
        if not old_post_game_path:
            os.makedirs(self.post_game_path)
        else:
            pass
        if not old_timeline_path:
            os.makedirs(self.timeline_path)
        else:
            pass
        if not old_picks_and_bans_path:
            os.makedirs(self.picks_and_bans_path)
        else:
            pass
        # analysis paths
        self.analysis_path = "./Analysis"
        self.draft_stats_path = self.analysis_path + "/Draft_Stats"
        self.draft_picks_and_bans_path = self.analysis_path + "/Draft_Stats/Picks_and_Bans"
        self.raw_stats_path = self.analysis_path + "/Raw_Stats"
        self.heatmaps_path = self.analysis_path + "/Heatmaps"
        self.histogram_path = self.analysis_path + "/Minimap_Histograms"
        # creates analysis paths is needed
        old_analysis_path = os.path.isdir(self.analysis_path)
        old_draft_stats_path = os.path.isdir(self.draft_stats_path)
        old_draft_picks_and_bans_path = os.path.isdir(self.draft_picks_and_bans_path)
        old_raw_stats_path = os.path.isdir(self.raw_stats_path)
        old_heatmaps_path = os.path.isdir(self.heatmaps_path)
        old_histogram_path = os.path.isdir(self.histogram_path)
        if not old_analysis_path:
            os.makedirs(self.analysis_path)
        else:
            pass
        if not old_draft_stats_path:
            os.makedirs(self.draft_stats_path)
        else:
            pass
        if not old_draft_picks_and_bans_path:
            os.makedirs(self.draft_picks_and_bans_path)
        else:
            pass
        if not old_raw_stats_path:
            os.makedirs(self.raw_stats_path)
        else:
            pass
        if not old_heatmaps_path:
            os.makedirs(self.heatmaps_path)
        else:
            pass
        if not old_histogram_path:
            os.makedirs(self.histogram_path)
        else:
            pass
        print(" Paths set correctly.")


    def update_games_data(self):
        
        ################################################
        # Login
        DRIVER_PATH = self.chromedriver_path
        options = Options()
        options.headless = True
        options.add_argument("--window-size=1920,1200")
        driver = webdriver.Chrome(options=options, executable_path=DRIVER_PATH)
        driver.get("https://euw.leagueoflegends.com/es-es/")
        # go to login page
        login_page = driver.find_element_by_xpath("//a[@data-riotbar-link-id='login']").click()
        time.sleep(1)
        # login
        login = driver.find_element_by_xpath("//input[@name='username']").send_keys(self.username)
        password = driver.find_element_by_xpath("//input[@name='password']").send_keys(self.password)
        submit = driver.find_element_by_xpath("//button[@type='submit']").click()
        time.sleep(1)


        # folders paths
        post_game_path = self.post_game_path + "/"
        timeline_path = self.timeline_path + "/"
        picks_and_bans_path = self.picks_and_bans_path + "/"

        #################################################
        # leaguepedia API request
        site = mwclient.Site('lol.gamepedia.com',path='/')
        tournament_name = self.tournament_name
        new_response = site.api('cargoquery',
                limit = 'max',
                join_on = "SG.OverviewPage=T.OverviewPage, SG.ScoreboardID_Wiki =PB.GameID_Wiki",
                tables = "ScoreboardGames=SG, Tournaments=T, PicksAndBansS7=PB",
                fields = "SG.Tournament, SG.ScoreboardID_Wiki , SG.DateTime_UTC, SG.Patch, SG.MatchHistory, SG.Winner, SG.Team1, SG.Team2, SG.Team1Picks, SG.Team2Picks, SG.Team1Bans, SG.Team2Bans, SG.Team1Players, SG.Team2Players, SG.OverviewPage, T.OverviewPage, T.Name, PB.GameID_Wiki, PB.Team1Ban1, PB.Team1Ban2, PB.Team1Ban3, PB.Team1Ban4, PB.Team1Ban5, PB.Team1Pick1, PB.Team1Pick2, PB.Team1Pick3, PB.Team1Pick4, PB.Team1Pick5, PB.Team2Ban1, PB.Team2Ban2, PB.Team2Ban3, PB.Team2Ban4, PB.Team2Ban5, PB.Team2Pick1, PB.Team2Pick2, PB.Team2Pick3, PB.Team2Pick4, PB.Team2Pick5",
                where = "T.name='" + str(tournament_name) + "'")

        #################################################
        # store the game's data in Games_Data/Post_Game
        try:
            for x in new_response['cargoquery']:
                time.sleep(2)
                # save json of the match
                # TO-DO: save the actual post game match data
                post_game = {}
                outfile = {}
                # queries the data
                post_game_link = "https://acs.leagueoflegends.com/v1/stats/game/" + str(x['title']['MatchHistory'].split('/')[5]) + "/" + str(x['title']['MatchHistory'].split('/')[6])
                print(post_game_link)
                post_game_data = driver.get(post_game_link)
                soup_2 = BeautifulSoup(driver.page_source,'html.parser')
                post_game_stats = json.loads(soup_2.find("body").text)

                post_game['info'] = []
                post_game['info'].append({
                    'tournament': tournament_name,
                    'blue_team': str(x['title']['Team1']),
                    'red_team': str(x['title']['Team2']),
                    'match_history': str(x['title']['MatchHistory']),
                    'date' : str(x['title']['DateTime UTC'][:10]),
                    'patch': str(x['title']['Patch'])
                })
                post_game['stats'] = post_game_stats
                # creates new path for tournament if it doesn't exists
                old_post_game_path = os.path.isdir(post_game_path + tournament_name)
                if not old_post_game_path:
                    os.makedirs((post_game_path + tournament_name))
                else:
                    pass
                    # print((post_game_path + tournament_name) ," folder already exists.")
                # saves the post game data json file
                with open(post_game_path +  tournament_name + '/' + str(x['title']['DateTime UTC'][:10]) + '_' + str(x['title']['Team1']) + '_' + str(x['title']['Team2']) + '.json','w') as outfile:
                    json.dump(post_game,outfile)
        except:
            print(" Game_Data/Post_Game failed storing the games data :(")
            
        #################################################
        # store the game's data in Games_Data/Timeline
        try:
            for x in new_response['cargoquery']:
                time.sleep(2)
                # save json of the match
                # TO-DO: save the actual timeline match data
                timeline = {}
                outfile = {}
                # queries the data
                timeline_link = "https://acs.leagueoflegends.com/v1/stats/game/" + str(x['title']['MatchHistory'].split('/')[5]) + "/" + str((x['title']['MatchHistory'].split('/')[6]).split('?')[0]) + "/timeline?" + str((x['title']['MatchHistory'].split('/')[6]).split('?')[1])
                print(timeline_link)
                timeline_data = driver.get(timeline_link)
                soup_1 = BeautifulSoup(driver.page_source,'html.parser')
                timeline_stats = json.loads(soup_1.find("body").text)

                timeline['info'] = []
                timeline['info'].append({
                    'tournament': tournament_name,
                    'blue_team': str(x['title']['Team1']),
                    'red_team': str(x['title']['Team2']),
                    'match_history': str(x['title']['MatchHistory']),
                    'date' : str(x['title']['DateTime UTC'][:10]),
                    'patch': str(x['title']['Patch'])
                })
                timeline['stats'] = timeline_stats
                # creates new path for tournament if it doesn't exists
                old_timeline_path = os.path.isdir(timeline_path + tournament_name)
                if not old_timeline_path:
                    os.makedirs((timeline_path + tournament_name))
                else:
                    pass
                    # print((timeline_path + tournament_name) ," folder already exists.")
                # saves the post game data json file
                with open(timeline_path +  tournament_name + '/' + str(x['title']['DateTime UTC'][:10]) + '_' + str(x['title']['Team1']) + '_' + str(x['title']['Team2']) + '.json','w') as outfile:
                    json.dump(timeline,outfile)
        except:
            print(" Game_Data/Timeline failed storing the games data :(")
            
        #################################################
        # store the game's drafts in Games_Data/Picks_and_Bans
        try:
            for x in new_response['cargoquery']:
                # save json of the draft
                picks_and_bans = {}
                # game info
                picks_and_bans['info'] = []
                picks_and_bans['info'].append({
                    'tournament': tournament_name,
                    'blue_team': str(x['title']['Team1']),
                    'red_team': str(x['title']['Team2']),
                    'winner': str(x['title']['Winner']),
                    'date' : str(x['title']['DateTime UTC'][:10]),
                    'patch': str(x['title']['Patch'])
                })
                # players
                picks_and_bans['players'] = []
                picks_and_bans['players'].append({
                    "blue":[
                        {
                            'player': x['title']['Team1Players'].split(',')[0],
                            'champion': x['title']['Team1Picks'].split(',')[0],
                            'role': 'Top'
                        },
                        {
                            'player': x['title']['Team1Players'].split(',')[1],
                            'champion': x['title']['Team1Picks'].split(',')[1],
                            'role': 'Jungle'
                        },
                        {
                            'player': x['title']['Team1Players'].split(',')[2],
                            'champion': x['title']['Team1Picks'].split(',')[2],
                            'role': 'Mid'
                        },
                        {
                            'player': x['title']['Team1Players'].split(',')[3],
                            'champion': x['title']['Team1Picks'].split(',')[3],
                            'role': 'Adc'
                        },
                        {
                            'player': x['title']['Team1Players'].split(',')[4],
                            'champion': x['title']['Team1Picks'].split(',')[4],
                            'role': 'Support'
                        }
                    ],
                    "red":[
                        {
                            'player': x['title']['Team2Players'].split(',')[0],
                            'champion': x['title']['Team2Picks'].split(',')[0],
                            'role': 'Top'
                        },
                        {
                            'player': x['title']['Team2Players'].split(',')[1],
                            'champion': x['title']['Team2Picks'].split(',')[1],
                            'role': 'Jungle'
                        },
                        {
                            'player': x['title']['Team2Players'].split(',')[2],
                            'champion': x['title']['Team2Picks'].split(',')[2],
                            'role': 'Mid'
                        },
                        {
                            'player': x['title']['Team2Players'].split(',')[3],
                            'champion': x['title']['Team2Picks'].split(',')[3],
                            'role': 'Adc'
                        },
                        {
                            'player': x['title']['Team2Players'].split(',')[4],
                            'champion': x['title']['Team2Picks'].split(',')[4],
                            'role': 'Support'
                        }
                    ]
                })
                # game draft
                picks_and_bans['draft'] = []
                picks_and_bans['draft'].append({
                    "picks":{
                        "blue":{
                            0: str(x['title']['Team1Pick1']),
                            1: x['title']['Team1Pick2'],
                            2: x['title']['Team1Pick3'],
                            3: x['title']['Team1Pick4'],
                            4: x['title']['Team1Pick5'],
                        },
                        "red":{
                            0: x['title']['Team2Pick1'],
                            1: x['title']['Team2Pick2'],
                            2: x['title']['Team2Pick3'],
                            3: x['title']['Team2Pick4'],
                            4: x['title']['Team2Pick5'],
                        }
                    },
                    "bans":{
                        "blue":{
                            0: x['title']['Team1Ban1'],
                            1: x['title']['Team1Ban2'],
                            2: x['title']['Team1Ban3'],
                            3: x['title']['Team1Ban4'],
                            4: x['title']['Team1Ban5'],
                        },
                        "red":{
                            0: x['title']['Team2Ban1'],
                            1: x['title']['Team2Ban2'],
                            2: x['title']['Team2Ban3'],
                            3: x['title']['Team2Ban4'],
                            4: x['title']['Team2Ban5'],
                        }
                    }
                })
                # creates new path for tournament if it doesn't exists
                old_picks_and_bans_path = os.path.isdir(picks_and_bans_path + tournament_name)
                if not old_picks_and_bans_path:
                    os.makedirs((picks_and_bans_path + tournament_name))
                else:
                    pass
                    # print((picks_and_bans_path + tournament_name) ," folder already exists.")
                # saves the post game data json file
                with open(picks_and_bans_path +  tournament_name + '/' + str(x['title']['DateTime UTC'][:10]) + '_' + str(x['title']['Team1']) + '_' + str(x['title']['Team2']) + '.json','w') as outfile:
                    json.dump(picks_and_bans,outfile)
        except:
            print(" Game_Data/Picks_and_Bans failed storing the games data :(")

        driver.quit()

    def get_competitive_stats(self):
                
        ########################################################################################################
        # leaguepedia API request
        site = mwclient.Site('lol.gamepedia.com',path='/')
        tournament_name = self.tournament_name
        new_response = site.api('cargoquery',
                limit = 'max',
                join_on = "SG.OverviewPage=T.OverviewPage, SG.ScoreboardID_Wiki =PB.GameID_Wiki",
                tables = "ScoreboardGames=SG, Tournaments=T, PicksAndBansS7=PB",
                fields = "SG.Tournament, SG.ScoreboardID_Wiki , SG.DateTime_UTC , SG.MatchHistory, SG.Winner, SG.Team1, SG.Team2, SG.OverviewPage, SG.Team1Players, SG.Team2Players, T.OverviewPage, T.Name, PB.GameID_Wiki",
                where = "T.name='" + str(tournament_name) + "'")

        ########################################################################################################
        # get champion ids
        champions = {}
        all_champions = requests.get("http://ddragon.leagueoflegends.com/cdn/10.21.1/data/en_US/champion.json").json()
        counter = 0
        for champ in all_champions['data']:
            champions.update({(all_champions['data'][champ]['key']): (all_champions['data'][champ]['id'])})
            counter += 1
        # --- finds the id of the requested champion ---
        def find_name(champ_id):
            for x in champions:
                if int(x) == int(champ_id):
                    return champions[x]



        ########################################################################################################
        # creates lists to fill in the stats
        game_id = {}
        date = {}
        team = {}
        team_rival = {}
        side = {}
        patch = {}
        role = {}
        player_name = {}
        champ = {}
        matchup = {}
        kda = {}
        kills = {}
        deaths = {}
        assists = {}
        kp = {}
        G_share = {}
        DMG_share = {}
        CSDat10 = {}
        GDat10 = {}
        XPat10 = {}
        game_time = {}
        result = {}

        row_counter = 0
        game_counter = 0


        timeline_path = self.timeline_path + "/"
        post_game_path = self.post_game_path + "/"
        ########################################################################################################
        # generate and save all the stats of every game in Analysis/Raw_Stats
        for x in new_response['cargoquery']:
            # read json
            with open(timeline_path +  tournament_name + '/' + str(x['title']['DateTime UTC'][:10]) + '_' + str(x['title']['Team1']) + '_' + str(x['title']['Team2']) + '.json') as f:
                data_timeline = json.load(f)
            with open(post_game_path +  tournament_name + '/' + str(x['title']['DateTime UTC'][:10]) + '_' + str(x['title']['Team1']) + '_' + str(x['title']['Team2']) + '.json') as p:
                data_post_game = json.load(p)
            
            # -------------------------- fills the workbook with game's stats --------------------------
            player_counter = 1
            col = 0
            

            for game in data_post_game['stats']['participants']:
                game_duration = 0
                totalTeamKills = 0
                totalTeamGold = 0
                totalTeamDamage = 0
                csBeforeMin10 = 0
                goldBeforeMin10 = 0
                xpBeforeMin10 = 0
                enemyGoldBeforeMin10 = 0
                enemyXpBeforeMin10 = 0
                enemyCsBeforeMin10 = 0

                # game id
                game_id[row_counter] = game_counter
        
                # week
                date[row_counter] = str(x['title']['DateTime UTC'][:10])

                if game['teamId'] == 100:
                    # team
                    team[row_counter] = str(x['title']['Team1'])
                    # side
                    side[row_counter] = 'blue'
                    # team rival
                    team_rival[row_counter] = str(x['title']['Team2'])
                else:
                    # team
                    team[row_counter] = str(x['title']['Team2'])
                    # side
                    side[row_counter] = 'red'
                    # team rival
                    team_rival[row_counter] = str(x['title']['Team1'])
                
                # patch
                patch[row_counter] = str(data_post_game['stats']['gameVersion'][:5])
                    
                # role
                if player_counter == 1 or player_counter == 6:
                    role[row_counter] = 'Top'
                    if player_counter == 1:
                        player_name[row_counter] = x['title']['Team1Players'].split(',')[0]
                    elif player_counter == 6:
                        player_name[row_counter] = x['title']['Team2Players'].split(',')[0]
                elif player_counter == 2 or player_counter == 7:
                    role[row_counter] = 'Jungle'
                    if player_counter == 2:
                        player_name[row_counter] = x['title']['Team1Players'].split(',')[1]
                    elif player_counter == 7:
                        player_name[row_counter] = x['title']['Team2Players'].split(',')[1]
                elif player_counter == 3 or player_counter == 8:
                    role[row_counter] = 'Mid'
                    if player_counter == 3:
                        player_name[row_counter] = x['title']['Team1Players'].split(',')[2]
                    elif player_counter == 8:
                        player_name[row_counter] = x['title']['Team2Players'].split(',')[2]
                elif player_counter == 4 or player_counter == 9:
                    role[row_counter] = 'Adc'
                    if player_counter == 4:
                        player_name[row_counter] = x['title']['Team1Players'].split(',')[3]
                    elif player_counter == 9:
                        player_name[row_counter] = x['title']['Team2Players'].split(',')[3]
                elif player_counter == 5 or player_counter == 10:
                    role[row_counter] = 'Support'
                    if player_counter == 5:
                        player_name[row_counter] = x['title']['Team1Players'].split(',')[4]
                    elif player_counter == 10:
                        player_name[row_counter] = x['title']['Team2Players'].split(',')[4]
                
                # champ
                champion = find_name(game['championId'])
                champ[row_counter] = champion

                # matchup
                for y in data_post_game['stats']['participants']:
                    if game['teamId'] == 100: 
                        if y['participantId'] == (player_counter+5):
                            champion = find_name(y['championId'])
                            matchup[row_counter] = champion
                    if game['teamId'] == 200:
                        if y['participantId'] == (player_counter-5):
                            champion = find_name(y['championId'])
                            matchup[row_counter] = champion
                
                
                # result
                if game['stats']['win'] == True:
                    result[row_counter] = 1
                else:
                    result[row_counter] = 0
                    
                
                if game['stats']['deaths'] > 0:
                    # kda
                    kda[row_counter] = ((game['stats']['kills']+game['stats']['assists'])/game['stats']['deaths'])
                else:
                    # kda
                    kda[row_counter] = (game['stats']['kills']+game['stats']['assists'])
                
                # kills
                kills[row_counter] = game['stats']['kills']
                # deaths
                deaths[row_counter] = game['stats']['deaths']
                # assists
                assists[row_counter] = game['stats']['assists']
                
                for players in data_post_game['stats']['participants']:
                    if players['teamId'] == game['teamId']:
                        # total teams stats
                        totalTeamKills += players['stats']['kills']
                        totalTeamGold += players['stats']['goldEarned']
                        totalTeamDamage += players['stats']['totalDamageDealt']
                        
                count = 0
                for frame in data_timeline['stats']['frames']:
                    if count == 10:
                        for participant in frame['participantFrames']:
                            if str(participant) == str(player_counter):
                                if player_counter < 6:   
                                    csBeforeMin10 = frame['participantFrames'][str(player_counter)]['minionsKilled'] + frame['participantFrames'][str(player_counter)]['jungleMinionsKilled']
                                    goldBeforeMin10 = frame['participantFrames'][str(player_counter)]['totalGold']
                                    xpBeforeMin10 = frame['participantFrames'][str(player_counter)]['xp']
                                    enemyCsBeforeMin10 = frame['participantFrames'][str(player_counter+5)]['minionsKilled'] + frame['participantFrames'][str(player_counter+5)]['jungleMinionsKilled']
                                    enemyGoldBeforeMin10 = frame['participantFrames'][str(player_counter+5)]['totalGold']
                                    enemyXpBeforeMin10 = frame['participantFrames'][str(player_counter+5)]['xp']
                                    
                                    break
                                if player_counter > 5:
                                    csBeforeMin10 = frame['participantFrames'][str(player_counter)]['minionsKilled'] + frame['participantFrames'][str(player_counter)]['jungleMinionsKilled']
                                    goldBeforeMin10 = frame['participantFrames'][str(player_counter)]['totalGold']
                                    xpBeforeMin10 = frame['participantFrames'][str(player_counter)]['xp']
                                    enemyCsBeforeMin10 = frame['participantFrames'][str(player_counter-5)]['minionsKilled'] + frame['participantFrames'][str(player_counter-5)]['jungleMinionsKilled']
                                    enemyGoldBeforeMin10 = frame['participantFrames'][str(player_counter-5)]['totalGold']
                                    enemyXpBeforeMin10 = frame['participantFrames'][str(player_counter-5)]['xp']
                                    break
                    count += 1
                            
                            

                if (game['stats']['kills']+game['stats']['assists']) != 0:
                    # kp
                    kp[row_counter] = (((game['stats']['kills']+game['stats']['assists'])/totalTeamKills)*100)
                else:
                    # kp
                    kp[row_counter] = 0
                
                # G%
                G_share[row_counter] = ((game['stats']['goldEarned']/totalTeamGold)*100)
                # DMG%
                DMG_share[row_counter] = ((game['stats']['totalDamageDealt']/totalTeamDamage)*100)
                # CSD@10
                CSDat10[row_counter] = (csBeforeMin10-enemyCsBeforeMin10)
                # GD@10
                GDat10[row_counter] = (goldBeforeMin10-enemyGoldBeforeMin10)
                # XPD@10
                XPat10[row_counter] = (xpBeforeMin10-enemyXpBeforeMin10)
                # GT
                game_time[row_counter] = str(data_post_game['stats']['gameDuration']/60)[:5]


                totalTeamKills = 0
                totalTeamGold = 0
                totalTeamDamage = 0
                csBeforeMin10 = 0
                goldBeforeMin10 = 0
                xpBeforeMin10 = 0
                enemyGoldBeforeMin10 = 0
                enemyXpBeforeMin10 = 0
                enemyCsBeforeMin10 = 0

                player_counter += 1
                row_counter += 1
            game_counter += 1


        ########################################################################################################
        # Excel
        first_word = 0
        tournament_name_file = ""
        for word in tournament_name.split(" "):
            if first_word == 0:
                tournament_name_file = word
                first_word = 1
            else:
                tournament_name_file = tournament_name_file + "_" + word

        raw_stats_path = self.raw_stats_path
        output_name = (raw_stats_path + "/" + tournament_name_file)
        workbook = xlsxwriter.Workbook(output_name + '.xlsx')
        worksheet = workbook.add_worksheet(tournament_name)


        worksheet.write(0,0,'game_id')
        worksheet.write(0,1,'date')
        worksheet.write(0,2,'team')
        worksheet.write(0,3,'team_rival')
        worksheet.write(0,4,'side')
        worksheet.write(0,5,'patch')
        worksheet.write(0,6,'role')
        worksheet.write(0,7,'player_name')
        worksheet.write(0,8,'champ')
        worksheet.write(0,9,'matchup')
        worksheet.write(0,10,'kda')
        worksheet.write(0,11,'kills')
        worksheet.write(0,12,'deaths')
        worksheet.write(0,13,'assists')
        worksheet.write(0,14,'kp')
        worksheet.write(0,15,'G%')
        worksheet.write(0,16,'DMG%')
        worksheet.write(0,17,'CSD@10')
        worksheet.write(0,18,'GD@10')
        worksheet.write(0,19,'XPD@10')
        worksheet.write(0,20,'game_time')
        worksheet.write(0,21,'Result')

        match_counter = 0
        row = 1
        col = 0
        for x in game_id:
            worksheet.write(row,col,game_id[match_counter])
            worksheet.write(row,col+1,date[match_counter])
            worksheet.write(row,col+2,team[match_counter])
            worksheet.write(row,col+3,team_rival[match_counter])
            worksheet.write(row,col+4,side[match_counter])
            worksheet.write(row,col+5,patch[match_counter])
            worksheet.write(row,col+6,role[match_counter])
            worksheet.write(row,col+7,player_name[match_counter])
            worksheet.write(row,col+8,champ[match_counter])
            worksheet.write(row,col+9,matchup[match_counter])
            worksheet.write(row,col+10,kda[match_counter])
            worksheet.write(row,col+11,kills[match_counter])
            worksheet.write(row,col+12,deaths[match_counter])
            worksheet.write(row,col+13,assists[match_counter])
            worksheet.write(row,col+14,kp[match_counter])
            worksheet.write(row,col+15,G_share[match_counter])
            worksheet.write(row,col+16,DMG_share[match_counter])
            worksheet.write(row,col+17,CSDat10[match_counter])
            worksheet.write(row,col+18,GDat10[match_counter])
            worksheet.write(row,col+19,XPat10[match_counter])
            worksheet.write(row,col+20,game_time[match_counter])
            worksheet.write(row,col+21,result[match_counter])
            row += 1
            match_counter += 1

        workbook.close()

    def draft_picks_and_bans(self):
        ########################################################################################################
        # leaguepedia API request
        site = mwclient.Site('lol.gamepedia.com',path='/')
        tournament_name = self.tournament_name
        new_response = site.api('cargoquery',
                limit = 'max',
                join_on = "SG.OverviewPage=T.OverviewPage, SG.ScoreboardID_Wiki =PB.GameID_Wiki",
                tables = "ScoreboardGames=SG, Tournaments=T, PicksAndBansS7=PB",
                fields = "SG.Tournament, SG.ScoreboardID_Wiki , SG.DateTime_UTC , SG.MatchHistory, SG.Winner, SG.Team1, SG.Team2, SG.Team1Picks, SG.Team2Picks, SG.Team1Bans, SG.Team2Bans, SG.Team1Players, SG.Team2Players, SG.OverviewPage, T.OverviewPage, T.Name, PB.GameID_Wiki, PB.Team1Ban1, PB.Team1Ban2, PB.Team1Ban3, PB.Team1Ban4, PB.Team1Ban5, PB.Team1Pick1, PB.Team1Pick2, PB.Team1Pick3, PB.Team1Pick4, PB.Team1Pick5, PB.Team2Ban1, PB.Team2Ban2, PB.Team2Ban3, PB.Team2Ban4, PB.Team2Ban5, PB.Team2Pick1, PB.Team2Pick2, PB.Team2Pick3, PB.Team2Pick4, PB.Team2Pick5",
                where = "T.name='" + str(tournament_name) + "'")



        ########################################################################################################
        # creates lists to fill in the draft picks/bans
        game_id = {}
        tournament = {}
        date = {}
        patch = {}
        side = {}
        team = {}
        result = {}

        ban_1 = {}
        ban_2 = {}
        ban_3 = {}
        ban_4 = {}
        ban_5 = {}
        pick_1 = {}
        pick_2 = {}
        pick_3 = {}
        pick_4 = {}
        pick_5 = {}

        picks_and_bans_path = self.picks_and_bans_path + "/"
        draft_path = self.draft_picks_and_bans_path + "/"
        ########################################################################################################
        # gets all the drafts
        counter = 0
        game_id_counter = 0
        for x in new_response['cargoquery']:
            # read json
            with open(picks_and_bans_path +  tournament_name + '/' + str(x['title']['DateTime UTC'][:10]) + '_' + str(x['title']['Team1']) + '_' + str(x['title']['Team2']) + '.json') as f:
                data_picks_and_bans = json.load(f)
            

            # ---------------------------------------------------- BLUE ----------------------------------------------------
            # game_id
            game_id[counter] = int(game_id_counter)
            # tournament
            tournament[counter] = str(data_picks_and_bans['info'][0]['tournament'])
            # date
            date[counter] = str(data_picks_and_bans['info'][0]['date'])
            # patch
            patch[counter] = str(data_picks_and_bans['info'][0]['patch'])
            # side
            side[counter] = "blue"
            # blue_team
            team[counter] = str(data_picks_and_bans['info'][0]['blue_team'])
            # result
            if data_picks_and_bans['info'][0]['winner'] == "1":
                result[counter] = 1
            elif data_picks_and_bans['info'][0]['winner'] == "2":
                result[counter] = 0   
            # ban_1
            ban_1[counter] = str(data_picks_and_bans['draft'][0]['bans']['blue']['0'])
            # ban__2
            ban_2[counter] = str(data_picks_and_bans['draft'][0]['bans']['blue']['1'])
            # ban_3
            ban_3[counter] = str(data_picks_and_bans['draft'][0]['bans']['blue']['2'])
            # ban_4
            ban_4[counter] = str(data_picks_and_bans['draft'][0]['bans']['blue']['3'])
            # ban_5
            ban_5[counter] = str(data_picks_and_bans['draft'][0]['bans']['blue']['4'])
            # pick_1
            pick_1[counter] = str(data_picks_and_bans['draft'][0]['picks']['blue']['0'])
            # pick_2
            pick_2[counter] = str(data_picks_and_bans['draft'][0]['picks']['blue']['1'])
            # pick_3
            pick_3[counter] = str(data_picks_and_bans['draft'][0]['picks']['blue']['2'])
            # pick_4
            pick_4[counter] = str(data_picks_and_bans['draft'][0]['picks']['blue']['3'])
            # pick_5
            pick_5[counter] = str(data_picks_and_bans['draft'][0]['picks']['blue']['4'])

            counter += 1

            # ---------------------------------------------------- RED ----------------------------------------------------
            # game_id
            game_id[counter] = int(game_id_counter)
            # tournament
            tournament[counter] = str(data_picks_and_bans['info'][0]['tournament'])
            # date
            date[counter] = str(data_picks_and_bans['info'][0]['date'])
            # patch
            patch[counter] = str(data_picks_and_bans['info'][0]['patch'])
            # side
            side[counter] = "red"
            # blue_team
            team[counter] = str(data_picks_and_bans['info'][0]['red_team'])
            # result
            if data_picks_and_bans['info'][0]['winner'] == "1":
                result[counter] = 0
            elif data_picks_and_bans['info'][0]['winner'] == "2":
                result[counter] = 1   
            # ban_1
            ban_1[counter] = str(data_picks_and_bans['draft'][0]['bans']['red']['0'])
            # ban__2
            ban_2[counter] = str(data_picks_and_bans['draft'][0]['bans']['red']['1'])
            # ban_3
            ban_3[counter] = str(data_picks_and_bans['draft'][0]['bans']['red']['2'])
            # ban_4
            ban_4[counter] = str(data_picks_and_bans['draft'][0]['bans']['red']['3'])
            # ban_5
            ban_5[counter] = str(data_picks_and_bans['draft'][0]['bans']['red']['4'])
            # pick_1
            pick_1[counter] = str(data_picks_and_bans['draft'][0]['picks']['red']['0'])
            # pick_2
            pick_2[counter] = str(data_picks_and_bans['draft'][0]['picks']['red']['1'])
            # pick_3
            pick_3[counter] = str(data_picks_and_bans['draft'][0]['picks']['red']['2'])
            # pick_4
            pick_4[counter] = str(data_picks_and_bans['draft'][0]['picks']['red']['3'])
            # pick_5
            pick_5[counter] = str(data_picks_and_bans['draft'][0]['picks']['red']['4'])

            counter += 1
            game_id_counter += 1


            
                            


        ########################################################################################################
        # Excel
        first_word = 0
        tournament_name_file = ""
        for word in tournament_name.split(" "):
            if first_word == 0:
                tournament_name_file = word
                first_word = 1
            else:
                tournament_name_file = tournament_name_file + "_" + word

        output_name = (draft_path + tournament_name_file)
        workbook = xlsxwriter.Workbook(output_name + '.xlsx')
        worksheet = workbook.add_worksheet(tournament_name)


        worksheet.write(0,0,'game_id')
        worksheet.write(0,1,'tournament')
        worksheet.write(0,2,'date')
        worksheet.write(0,3,'patch')
        worksheet.write(0,4,'side')
        worksheet.write(0,5,'team')
        worksheet.write(0,6,'result')
        worksheet.write(0,7,'ban_1')
        worksheet.write(0,8,'ban_2')
        worksheet.write(0,9,'ban_3')
        worksheet.write(0,10,'ban_4')
        worksheet.write(0,11,'ban_5')
        worksheet.write(0,12,'pick_1')
        worksheet.write(0,13,'pick_2')
        worksheet.write(0,14,'pick_3')
        worksheet.write(0,15,'pick_4')
        worksheet.write(0,16,'pick_5')


        match_counter = 0
        row = 1
        col = 0
        for x in game_id:
            worksheet.write(row,col,game_id[match_counter])
            worksheet.write(row,col+1,tournament[match_counter])
            worksheet.write(row,col+2,date[match_counter])
            worksheet.write(row,col+3,patch[match_counter])
            worksheet.write(row,col+4,side[match_counter])
            worksheet.write(row,col+5,team[match_counter])
            worksheet.write(row,col+6,result[match_counter])
            worksheet.write(row,col+7,ban_1[match_counter])
            worksheet.write(row,col+8,ban_2[match_counter])
            worksheet.write(row,col+9,ban_3[match_counter])
            worksheet.write(row,col+10,ban_4[match_counter])
            worksheet.write(row,col+11,ban_5[match_counter])
            worksheet.write(row,col+12,pick_1[match_counter])
            worksheet.write(row,col+13,pick_2[match_counter])
            worksheet.write(row,col+14,pick_3[match_counter])
            worksheet.write(row,col+15,pick_4[match_counter])
            worksheet.write(row,col+16,pick_5[match_counter])
            row += 1
            match_counter += 1

        workbook.close()