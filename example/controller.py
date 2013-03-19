#!/usr/bin/python
# coding: utf-8

import sys
sys.path.append("../")
from dwindow import DWindowUtiles

import libs
import config

class DWindowUtiles(DWindowUtiles):
    def __init__(self):
        process_name = u"디아블로 III"
        super(DWindowUtiles, self).__init__(process_name)


class Items:
    images = ['2_amulet.bmp',
              '2_evil_water.bmp',
              '2_ring.bmp',
              #'2_runesward.bmp',
              'runesward2.bmp',
              '2_domyon.bmp']

class Skills:
    war_throw2 = {'1': ['1_overpower.bmp', 2],
                 '2': ['1_smenu_weaponThrow.bmp', 2],
                 '3': ['1_smenu_wrath.bmp', 2],
                 '4': ['1_smenu_rage.bmp', 1],
                 'a': ['1_smenu_call.bmp', 2],
                 'b': ['1_smenu_earthquak.bmp', 3]}

    war_jump = {'1': ['1_smenu_leaf.bmp', 4],
                 '2': ['1_smenu_shout.bmp', 4],
                 '3': ['1_smenu_wrath.bmp', 2],
                 '4': ['1_smenu_rage.bmp', 1],
                 'a': ['1_smenu_cleave.bmp', 4],
                 'b': ['1_smenu_earthquak.bmp', 3]}
    
    menu_location = {'1': [232, 570], '2': [266,571], '3': [305,572], '4': [342,572], 'a': [380,574], 'b': [417,573]}
    menu_sub_location = [[200,243], [274,239], [360,239], [442, 239], [519,239], [598,241]]
    menu_next_button_location = [619, 112]

class D3ControllerBase(DWindowUtiles):
    def click(self, x, y):
        self.sendMouseLeftClick(x, y)
        
    def move(self, x, y):
        # change to middle button
        self.sendMouseMiddleClick(x, y)
        
    def select(self, x, y):
        self.setCursorPos(x, y)
        libs.sleep(300)
        self.clickMouseLeftButton()

    def skill1(self): self.sendKey("1")
    def skill2(self): self.sendKey("2")
    def skill3(self): self.sendKey("3")
    def skill4(self): self.sendKey("4")
    def skilla(self):
        self.sendMouseLeftClickEx(1, 1, shift=True)
    def skillb(self):
        self.sendMouseRightClickEx(1, 1, shift=True)

    def swingStandingLefta(self, count, sleep):
        #self.setCursorPos(300, 300)
        self.sendMouseLeftClickEx(300, 300, count=count, shift=True, sleep=sleep)
    def swingStandingRighta(self, count, sleep):
        #self.setCursorPos(500, 300)
        self.sendMouseLeftClickEx(500, 300, count=count, shift=True, sleep=sleep)

    def pickItems(self):
        self.sendAlt()
        libs.sleep(600)
        # legend
        #(109, 182-192, 150-161) #more tight
        #(109, 179-195, 123-171) #problem
        for a in range(1):
            self.pickItem([109, 182, 150], [109, 192, 161], 2)
            self.pickItem([60, 245, 107], [62, 255, 255], 2)            
            libs.sleep(300)
        
        # rare
        #(90-90, 234-254, 136-241)
        # libs.sleep(500)
        # self.pickItem([90, 234, 136], [90, 254, 241])

        self.pickItemImages()

    def pickItem(self, min_hsv, max_hsv, min_pixels):
        self.setCursorPos(10, 10)
        contours = self.findHSVEX(min_hsv, max_hsv, min_pixels)
        #print contours
        try:
            x, y = contours[0][0][0]
            print "PickItem: ", x, y
        except:
            return

        self.setCursorPos(x, y)
        libs.sleep(400)
        self.clickMouseLeftButton()
        libs.sleep(1800)


    def pickItemImages(self):
        for img in Items.images:
            libs.sleep(400)
            img = config.imgs_dir + img

            self.pickItemImage(img)
        
    def pickItemImage(self, img):
        try:
            maxLoc = self.isImageAtWindow(img)
            self.setCursorPos(*maxLoc)
            libs.sleep(400)
            self.clickMouseLeftButton()
            libs.sleep(1800)
        except:
            pass
                

    def setSkills(self, skills):
        for skill in skills:
            menu_loc = Skills.menu_location[skill]
            img, sub_skill_loc = skills[skill]
            img = config.imgs_dir + img

            # Open skill menu
            self.sendMouseRightClick(*menu_loc)
            libs.sleep(400)

            # Set skill
            self._setSkills(img, sub_skill_loc)

    def _setSkills(self, img, sub_skill_loc):
        for a in range(10):
            try:
                libs.sleep(800)
                # (229, 69), (577, 199)
                maxloc = self.isImageAtRegion(img, 229, 69, 350, 130)
                self.click(*maxloc)
                libs.sleep(400)                
                self.click(*Skills.menu_sub_location[sub_skill_loc])
                libs.sleep(400)
                self.click(338, 455)
                libs.sleep(400)
                return
            except:
                # Next page
                self.click(*Skills.menu_next_button_location)
                libs.sleep(400)
                continue

            # maxloc = self.isImageAtRegion(img, 225, 75, 363, 106)
            # self.click(*maxloc)
            # libs.sleep(400)                
            # self.click(*Skills.menu_sub_location[sub_skill_loc])
            # libs.sleep(400)
            # return
            

class D3BackgroundController(D3ControllerBase):
    pass

class D3ForegroundController(D3ControllerBase):
    pass

class D3Controller:

    def __init__(self, controller):
        self.controller = controller()
        self.__set_dict()

    def __set_dict(self):
        for attr in dir(self.controller):
            self.__dict__[attr] = getattr(self.controller, attr)
