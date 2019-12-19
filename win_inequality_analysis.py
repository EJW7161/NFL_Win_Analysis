#Call final_analysis() for final conclusions

import xlrd


#The excel sheet with stats needed
loc = ("C:\\Users\\ethan\\Projects\\Excel\\NFL Win Inequality\\Super Bowl Era Best and Worst Teams.xlsx")


#To open Workbook
wb = xlrd.open_workbook(loc)


#Making sheet from wb
global sheet
sheet = wb.sheet_by_index(0)

def reset():
    #All variables needed
    global year
    year = 1966 #1970 for modern era, 1966 for super bowl era
    
    global current_year
    current_year = 2019 #sheet's last year plus 1
    
    global year_col
    year_col = 0
    
    global mvp_col
    mvp_col = 0
    
    global best_number_teams_col
    best_number_teams_col = 2
    
    global worst_number_teams_col
    worst_number_teams_col = 9
    
    global best_team1_col
    best_team1_col = 2
    
    global best_team2_col
    best_team2_col = 3
    
    global best_team3_col
    best_team3_col = 8
    
    global best_team4_col
    best_team4_col = 1
    
    global best_team5_col
    best_team5_col = 15
    
    global best_team_wins_col
    best_team_wins_col = 4
    
    global best_team_losses_col
    best_team_losses_col = 5
    
    global best_team_ties_col
    best_team_ties_col = 6
    
    global best_team_pct_col
    best_team_pct_col = 7
    
    global worst_team1_col
    worst_team1_col = 9
    
    global worst_team2_col
    worst_team2_col = 10
    
    global worst_team3_col
    worst_team3_col = 15
    
    global worst_team_wins_col
    worst_team_wins_col = 11
    
    global worst_team_losses_col
    worst_team_losses_col = 12
    
    global worst_team_ties_col
    worst_team_ties_col = 13
    
    global worst_team_pct_col
    worst_team_pct_col = 14


#Functions to be called for year specific stats
def get_inequality_pct(year):
    row = 3 * (year-1965) - 1
    best = float(sheet.cell_value(row, best_team_pct_col))
    worst = float(sheet.cell_value(row, worst_team_pct_col))
    inequality_percent = best - worst
    return inequality_percent

def get_diff_wins(year):
    row = 3 * (year-1965) - 1
    best = int(sheet.cell_value(row, best_team_wins_col))
    worst = int(sheet.cell_value(row, worst_team_wins_col))
    diff_in_wins = best - worst
    return diff_in_wins
    
def get_best_teams(year):
    row = 3 * (year-1965) - 1
    number_of_best_teams = int(sheet.cell_value(row + 2,best_number_teams_col))
    if number_of_best_teams == 1:
        best_teams = [sheet.cell_value(row,best_team1_col)]
    elif number_of_best_teams == 2:
        best_teams = [sheet.cell_value(row,best_team1_col),
        sheet.cell_value(row,best_team2_col)]
    elif number_of_best_teams == 3:
        best_teams = [sheet.cell_value(row,best_team1_col),
        sheet.cell_value(row,best_team2_col),
        sheet.cell_value(row,best_team3_col)]
    elif number_of_best_teams == 4:
        best_teams = [sheet.cell_value(row,best_team1_col),
        sheet.cell_value(row,best_team2_col),
        sheet.cell_value(row,best_team3_col),
        sheet.cell_value(row,best_team4_col)]
    elif number_of_best_teams == 5:
        best_teams = [sheet.cell_value(row,best_team1_col),
        sheet.cell_value(row,best_team2_col),
        sheet.cell_value(row,best_team3_col),
        sheet.cell_value(row,best_team4_col),
        sheet.cell_value(row,best_team5_col)]
    return best_teams
    
def get_worst_teams(year):
    row = 3 * (year-1965) - 1
    number_of_worst_teams = int(sheet.cell_value(row + 2,worst_number_teams_col))
    if number_of_worst_teams == 1:
        worst_teams = [sheet.cell_value(row,worst_team1_col)]
    elif number_of_worst_teams == 2:
        worst_teams = [sheet.cell_value(row,worst_team1_col),
        sheet.cell_value(row,worst_team2_col)]
    elif number_of_worst_teams == 3:
        worst_teams = [sheet.cell_value(row,worst_team1_col),
        sheet.cell_value(row,worst_team2_col),
        sheet.cell_value(row,worst_team3_col)]
    return worst_teams

def get_superbowl_appearences(year):
    row = 3 * (year-1965)
    number_of_best_teams = int(sheet.cell_value(row + 1,best_number_teams_col))
    superbowl_appearences = 0
    if number_of_best_teams == 1:
        if sheet.cell_value(row,best_team1_col) != 'DNM SB':
            superbowl_appearences += 1
    elif number_of_best_teams == 2:
        if sheet.cell_value(row,best_team1_col) != 'DNM SB':
            superbowl_appearences += 1
        if sheet.cell_value(row,best_team2_col) != 'DNM SB':
            superbowl_appearences += 1
    elif number_of_best_teams == 3:
        if sheet.cell_value(row,best_team1_col) != 'DNM SB':
            superbowl_appearences += 1
        if sheet.cell_value(row,best_team2_col) != 'DNM SB':
            superbowl_appearences += 1
        if sheet.cell_value(row,best_team3_col) != 'DNM SB':
            superbowl_appearences += 1
    elif number_of_best_teams == 4:
        if sheet.cell_value(row,best_team1_col) != 'DNM SB':
            superbowl_appearences += 1
        if sheet.cell_value(row,best_team2_col) != 'DNM SB':
            superbowl_appearences += 1
        if sheet.cell_value(row,best_team3_col) != 'DNM SB':
            superbowl_appearences += 1
        if sheet.cell_value(row,best_team4_col) != 'DNM SB':
            superbowl_appearences += 1
    elif number_of_best_teams == 5:
        if sheet.cell_value(row,best_team1_col) != 'DNM SB':
            superbowl_appearences += 1
        if sheet.cell_value(row,best_team2_col) != 'DNM SB':
            superbowl_appearences += 1
        if sheet.cell_value(row,best_team3_col) != 'DNM SB':
            superbowl_appearences += 1
        if sheet.cell_value(row,best_team4_col) != 'DNM SB':
            superbowl_appearences += 1
        if sheet.cell_value(row,best_team5_col) != 'DNM SB':
            superbowl_appearences += 1
    return superbowl_appearences
    
def get_superbowl_wins(year):
    row = 3 * (year-1965)
    number_of_best_teams = sheet.cell_value(row + 1,best_number_teams_col)
    superbowl_wins = 0
    if number_of_best_teams == 1:
        if sheet.cell_value(row,best_team1_col) == 'Won SB':
            superbowl_wins += 1
    elif number_of_best_teams == 2:
        if sheet.cell_value(row,best_team1_col) == 'Won SB':
            superbowl_wins += 1
        if sheet.cell_value(row,best_team2_col) == 'Won SB':
            superbowl_wins += 1
    elif number_of_best_teams == 3:
        if sheet.cell_value(row,best_team1_col) == 'Won SB':
            superbowl_wins += 1
        if sheet.cell_value(row,best_team2_col) == 'Won SB':
            superbowl_wins += 1
        if sheet.cell_value(row,best_team3_col) == 'Won SB':
            superbowl_wins += 1
    elif number_of_best_teams == 4:
        if sheet.cell_value(row,best_team1_col) == 'Won SB':
            superbowl_wins += 1
        if sheet.cell_value(row,best_team2_col) == 'Won SB':
            superbowl_wins += 1
        if sheet.cell_value(row,best_team3_col) == 'Won SB':
            superbowl_wins += 1
        if sheet.cell_value(row,best_team4_col) == 'Won SB':
            superbowl_wins += 1
    elif number_of_best_teams == 5:
        if sheet.cell_value(row,best_team1_col) == 'Won SB':
            superbowl_wins += 1
        if sheet.cell_value(row,best_team2_col) == 'Won SB':
            superbowl_wins += 1
        if sheet.cell_value(row,best_team3_col) == 'Won SB':
            superbowl_wins += 1
        if sheet.cell_value(row,best_team4_col) == 'Won SB':
            superbowl_wins += 1
        if sheet.cell_value(row,best_team5_col) == 'Won SB':
            superbowl_wins += 1
    return superbowl_wins
    
def get_number_of_best_teams(year):
    row = 3 * (year-1965) + 1
    number_of_best_teams = int(sheet.cell_value(row,best_number_teams_col))
    return number_of_best_teams
    
def get_number_of_worst_teams(year):
    row = 3 * (year-1965) + 1
    number_of_worst_teams = int(sheet.cell_value(row,worst_number_teams_col))
    return number_of_worst_teams

#Functions to be called for historical stats    
def inequality_pcts():
    reset()
    global year
    global current_year
    global inequality_pcts
    inequality_pcts = []
    while year < current_year:
        inequality_pcts.append(get_inequality_pct(year))
        year += 1
    return inequality_pcts

def diff_wins():
    reset()
    global year
    global current_year
    global diff_wins
    diff_wins = []
    while year < current_year:
        diff_wins.append(get_diff_wins(year))
        year += 1
    return diff_wins

def best_teams():
    reset()
    global year
    global current_year
    global best_teams
    best_teams = []
    while year < current_year:
        year_best = get_best_teams(year)
        for team in year_best:
            best_teams.append(team)
        year += 1
    return best_teams
    
def worst_teams():
    reset()
    global year
    global current_year
    global worst_teams
    worst_teams = []
    while year < current_year:
        year_worst = get_worst_teams(year)
        for team in year_worst:
            worst_teams.append(team)
        year += 1
    return worst_teams
    
def total_superbowl_appearences():
    reset()
    global year
    global current_year
    global total_superbowl_appearences
    total_superbowl_appearences = 0
    while year < current_year:
        #only considering years with solo win leader as teams that lose or dont 
        #make super bowl likely fall to other teams with best record, messes up 
        #stats
        if get_number_of_best_teams(year) == 1:
            total_superbowl_appearences += get_superbowl_appearences(year)
        year += 1
    return total_superbowl_appearences
    
def total_superbowl_wins():
    reset()
    global year
    global current_year
    global total_superbowl_wins
    total_superbowl_wins = 0
    while year < current_year:
        if get_number_of_best_teams(year) == 1:
            total_superbowl_wins += get_superbowl_wins(year)
        year += 1
    return total_superbowl_wins
    
def total_number_of_best_teams():
    reset()
    global year
    global current_year
    global total_number_of_best_teams
    total_number_of_best_teams = 0
    while year < current_year:
        if get_number_of_best_teams(year) == 1:
            total_number_of_best_teams += get_number_of_best_teams(year)
        year += 1
    return total_number_of_best_teams
    
def super_bowl_analysis():
    appearences = total_superbowl_appearences()
    wins = total_superbowl_wins()
    number_teams = total_number_of_best_teams()
    super_bowl_appearence_rate = 100*appearences/number_teams
    super_bowl_win_rate = 100*wins/number_teams
    print('The solo NFL leader in record appears in the super bowl ' +
    str(super_bowl_appearence_rate) + '% of the time')
    print('The solo NFL leader in record wins the super bowl ' +
    str(super_bowl_win_rate) + '% of the time')
        
    
def total_number_of_worst_teams():
    reset()
    global year
    global current_year
    global total_number_of_worst_teams
    total_number_of_worst_teams = 0
    while year < current_year:
        total_number_of_worst_teams += get_number_of_worst_teams(year)
        year += 1
    return total_number_of_worst_teams
    
def percentage_inequality_retriever():
    reset()
    inequality_percentages = inequality_pcts()
    max_inequality_indeces = inequality_percentages.index(max(inequality_percentages))
    min_inequality_indeces = inequality_percentages.index(min(inequality_percentages))
    
    if type(max_inequality_indeces) == list:
        for index in len(max_inequality_indeces):
            max_inequality_years = []
            max_inequality_years.append(1966 + index)
    else: 
        max_inequality_years = 1966 + max_inequality_indeces

    if type(min_inequality_indeces) == list:
        for index in len(min_inequality_indeces):
            min_inequality_years = []
            min_inequality_years.append(1966 + index)
    else:
        min_inequality_years = 1966 + min_inequality_indeces

    print("The year(s) with the biggest difference in win percentage " +
    "between the best and worst team(s) was/were in " +
    str(max_inequality_years) + " with a difference of " 
    + str(max(inequality_percentages)))
    
    print("The year(s) with the smallest difference in win percentage " +
    "between the best and worst team(s) was/were in " +
    str(min_inequality_years) + " with a difference of " 
    + str(min(inequality_percentages)))
    
def win_inequality_retriever():
    reset()
    inequality_wins = diff_wins()
    max_inequality_indeces = inequality_wins.index(max(inequality_wins))
    min_inequality_indeces = inequality_wins.index(min(inequality_wins))
    
    if type(max_inequality_indeces) == list:
        for index in len(max_inequality_indeces):
            max_inequality_years = []
            max_inequality_years.append(1966 + index)
    else: 
        max_inequality_years = 1966 + max_inequality_indeces

    if type(min_inequality_indeces) == list:
        for index in len(min_inequality_indeces):
            min_inequality_years = []
            min_inequality_years.append(1966 + index)
    else:
        min_inequality_years = 1966 + min_inequality_indeces

    print("The year(s) with the biggest difference in wins " +
    "between the best and worst team(s) was/were in " +
    str(max_inequality_years) + " with a difference of " 
    + str(max(inequality_wins)))
    
    print("The year(s) with the smallest difference in wins " +
    "between the best and worst team(s) was/were in " +
    str(min_inequality_years) + " with a difference of " 
    + str(min(inequality_wins)))

global teams
teams = ['ARI Cardinals','ATL Falcons','BAL Colts','BAL Ravens',
'BOS Patriots','BUF Bills','CAR Panthers','CIN Bengals','CHI Bears',
'CLE Browns','DAL Cowboys','DEN Broncos','DET Lions','GB Packers',
'HOU Oilers','HOU Texans','IND Colts','JAX Jaguars','KC Chiefs',
'SD Chargers','LA Chargers','LA Rams','STL Rams','MIA Dolphins',
'MIN Vikings','NE Patriots','NO Saints','NY Giants','NY Jets','LA Raiders',
'OAK Raiders','PHI Eagles','PIT Steelers','SF 49ers','SEA Seahawks',
'TB Bucaneers','TEN Titans','WAS Redskins']

#Franchise IND Colts encompasses teams  BAL Colts and IND Colts
#Franchise NE Patriots encompasses teams NE Patriots and BOS Patriots
#Franchise TEN Titans encompasses teams HOU Oilers and TEN Titans
#Franchise LA Chargers encompasses teams SD Chargers and LA Chargers
#Franchise LA Rams encompasses teams STL Rams and LA Rams
#Franchise OAK Raiders encompasses teams LA Raiders and OAK Raiders

global franchises
franchises = ['ARI Cardinals','ATL Falcons','BAL Ravens',
'BUF Bills','CAR Panthers','CIN Bengals','CHI Bears',
'CLE Browns','DAL Cowboys','DEN Broncos','DET Lions','GB Packers',
'HOU Texans','IND Colts','JAX Jaguars','KC Chiefs',
'LA Chargers','LA Rams','MIA Dolphins',
'MIN Vikings','NE Patriots','NO Saints','NY Giants','NY Jets',
'OAK Raiders','PHI Eagles','PIT Steelers','SF 49ers','SEA Seahawks',
'TB Bucaneers','TEN Titans','WAS Redskins']

def worst_appearences_per_team(team, all_worst_teams):
    worst_appearences = 0
    for team_in_question in all_worst_teams:
        if team == team_in_question:
            worst_appearences += 1
    return worst_appearences

def worst_appearences():
    reset()
    global all_worst_teams
    all_worst_teams = worst_teams()
    global teams_and_appearences
    teams_and_appearences = []
    
    for team in teams:
        appearences = worst_appearences_per_team(team, all_worst_teams)
        teams_and_appearences.append([team, appearences])
        
    return teams_and_appearences
    print (teams_and_appearences)
    
def best_appearences_per_team(team, all_best_teams):
    best_appearences = 0
    for team_in_question in all_best_teams:
        if team == team_in_question:
            best_appearences += 1
    return best_appearences

def best_appearences():
    reset()
    global all_best_teams
    all_best_teams = best_teams()
    global teams_and_appearences
    teams_and_appearences = []
    
    for team in teams:
        appearences = best_appearences_per_team(team, all_best_teams)
        teams_and_appearences.append([team, appearences])
        
    return teams_and_appearences
    
def polarized_outliers():
    reset()
    b_appearences = []
    w_appearences = []
    
    best_teams_appearences = best_appearences()
    worst_teams_appearences = worst_appearences()
    
    for i in best_teams_appearences:
        b_appearences.append(i[1])

    polarized_best_team = best_teams_appearences[b_appearences.index(max(b_appearences))][0]
    print('The team that appeared the most as the team with the most wins' + 
        ' is/are the ' + polarized_best_team)
        
    for i in worst_teams_appearences:
        w_appearences.append(i[1])

    polarized_worst_team = worst_teams_appearences[w_appearences.index(max(w_appearences))][0]
    print('The team that appeared the most as the team with the least wins' + 
        ' is/are the ' + polarized_worst_team)
    
    total_appearences = []
    j = 0
    while j < len(best_teams_appearences):
        total_appearences.append(best_teams_appearences[j][1] + worst_teams_appearences[j][1])
        j += 1
        
    most_polarized_team_index = total_appearences.index(max(total_appearences))
    least_polarized_team_index = total_appearences.index(min(total_appearences))
    
    most_polarized_team = best_teams_appearences[most_polarized_team_index][0]
    most_appearences = total_appearences[most_polarized_team_index]
    least_polarized_team = best_teams_appearences[least_polarized_team_index][0]
    least_appearences = total_appearences[least_polarized_team_index]
    
    print('The team with the most appearences combined as the best and worst' +
    ' team in a given season is/are the ' + most_polarized_team + ' with ' +
    str(most_appearences))
    print('The team with the least appearences combined as the best and worst' +
    ' team in a given season is/are the ' + least_polarized_team + ' with ' +
    str(least_appearences))
    
def get_mvp(year):
    row = 3 * (year-1965)
    if sheet.cell_value(row,mvp_col) == 'MVP':
        mvp = True
    else:
        mvp = False
    return mvp
    
def mvp_analysis():
    reset()
    global year
    global current_year
    mvp_counter = 0
    while year < current_year:
        row = 3 * (year-1965)
        if sheet.cell_value(row,mvp_col) == 'MVP':
            mvp_counter += 1
        year += 1
    
    reset()
    
    print('The MVP is on the team with the best record in the NFL ' + 
        'or the team tied for the best record in the NFL ' +
        str(float(100*mvp_counter/(current_year-year))) + '% of the time')
   
def final_analysis():
    super_bowl_analysis()
    mvp_analysis()
    percentage_inequality_retriever()
    win_inequality_retriever()
    polarized_outliers()
