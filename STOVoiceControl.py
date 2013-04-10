import win32com.client
if win32com.client.gencache.is_readonly == True:

    #allow gencache to create the cached wrapper objects
    win32com.client.gencache.is_readonly = False

    # under p2exe the call in gencache to __init__() does not happen
    # so we use Rebuild() to force the creation of the gen_py folder
    win32com.client.gencache.Rebuild()

    # NB You must ensure that the python...\win32com.client.gen_py dir does not exist
    # to allow creation of the cache in %temp%

from dragonfly import CompoundRule, Grammar, MappingRule, Key, Function, Mouse
from dragonfly.grammar.elements import Dictation
from dragonfly.engines.engine import get_sapi5_engine

import pythoncom
import win32api, win32con
from win32com.client import constants
import win32api, win32con
from time import sleep
import json

defaultmapping = {
        "KeyBindings": {
                "maximum warp": "g",
                "go to maximum warp": "g",
                "full impulse": "g",
                "go to full impulse": "g",
                "full stop": "r",
                "take us into orbit": "f",
                "explore planet": "f",
                "scan the area": "v",
                "target enemy": "g",
                "scan anomaly": "f",
                "power to all shields": "del",
                "power to forward shields": "up",
                "power to rear shields": "down",
                "power to aft shields": "down",
                "power to starboard shields": "left",
                "power to left shields": "left",
                "power to right shields": "right",
                "fire phasers": "space",
                "fire all weapons": "a-space",
            },
        "MouseBindings": {
            }
    }

def echo(val):
    print "Command executed: ", val
    
def ForceLShiftDown(val):
    win32api.keybd_event(win32con.VK_LSHIFT, 0xaa, 0, 0)

def ForceLShiftUp():
    win32api.keybd_event(win32con.VK_LSHIFT, 0xaa, win32con.KEYEVENTF_KEYUP, 0)
       
def MappingFromFile():
    try:
        with open("config.txt", 'r') as infile:
            config = json.load(infile)
    except Exception as e:
        print "Creating default mapping!"
        print "Mapping is %s" % defaultmapping
        with open("config.txt", "w") as outfile:
            json.dump(defaultmapping, outfile, indent=4)
            outfile.flush()
        config = defaultmapping
    
    mapping = {}
    for key in config["KeyBindings"].keys(): 
        print "Found instruction: %s" % key
        mapping[key] = Key(config["KeyBindings"][key]) + Function(echo, val=key)
    for key in config["MouseBindings"].keys(): 
        mapping[key] = Mouse(config["MouseBindings"][key]) + Function(echo, val=key)
    return mapping
    
class StarshipRule(MappingRule):
    mapping  = MappingFromFile()
    
# Create a grammar which contains and loads the command rule.
grammar = Grammar("starship grammar")                          
rule = StarshipRule()
grammar.add_rule(rule)
grammar.load()

while True:
    pythoncom.PumpWaitingMessages()
    sleep(.1)