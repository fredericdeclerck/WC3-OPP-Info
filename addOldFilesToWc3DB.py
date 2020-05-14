# -*- coding: utf-8 -*-
"""
Created on Thu May  7 17:01:12 2020

@author: fred
"""
#this file is meant to add the information from
#past replays into the excel database
#the information is in the file name
#the scheme here is that most files have this structure
# "opponent_name strat.w3g"
#hence we will populate a file with 2 columns "Opponent" and "strat" from the database

import glob
import os
import pandas as pd

#Constants
DATA_FILE = "M:/Installed Apps/blizz/Warcraft III/oldReplaysInfo.xlsx"
REPLAYS_FOLDER = "C:/Users/fred/Documents/Warcraft III/Replay/Autosaved/Multiplayer"
FILE_STRUCT = "* *.w3g"
DELIMITER = " "
OPP_NAME_COL = "Opponent"
STRAT_COL = "Strat"

#################################
#
def write_info(opps, strats):
  data = {OPP_NAME_COL: opps, STRAT_COL: strats}
  df = pd.DataFrame(data = data)
  df.to_excel(DATA_FILE)

#################################
#
def retrieve_info():
  os.chdir(REPLAYS_FOLDER)
  player_list = []
  strat_list = []
  separator_position = 0

  for file in glob.glob(FILE_STRUCT):
      #print(file)
      separator_position = file.find(DELIMITER)
      player_list.append(file[0:separator_position])
      strat_list.append(file[separator_position+1: len(file)])

  return player_list, strat_list

############### Main
x, y = retrieve_info()
write_info(x, y)
