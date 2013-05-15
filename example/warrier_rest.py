#!/usr/bin/python
# coding: utf-8

import sys
sys.path.append("../")

import win32gui
import win32api
import win32con

import config
import libs
from controller import D3BackgroundController, D3ForegroundController, D3Controller, DWindowUtiles, Skills

def MAKELONG(l,h):
    return ((h & 0xFFFF) << 16) | (l & 0xFFFF)
  
MAKELPARAM=MAKELONG

class Town(object):
    '''
    >>> controller = D3Controller(D3ForegroundController)
    >>> town = Town(controller)
    >>> #x = 709
    >>> #y = 62
    >>> #town.ctl.setCursorPos(x, y)
    >>> #color = town.ctl.getPixel(x, y)
    >>> #color
    >>> #libs.rgb_to_hex(color)
    >>> town.isInTown()
    True

    >>> town.test_action()

    >>> town.isInMenu() #doctest: +SKIP

    '''
    def __init__(self, controller):
        self.ctl = controller
        self.window = self.ctl.window

    def test_action(self):
        # self._into_start_point()
        # self.waitTown()
        # libs.rSleep(1300)
        # self.goIntoWarrior()

        # current_hwnd = win32gui.GetForegroundWindow()
        # current_cursor_pos = win32gui.GetCursorPos()
        
        # x = self.window.left + 289
        # y = self.window.top + 135
        # win32gui.SetForegroundWindow(self.window.hwnd)
        # win32api.SetCursorPos((x, y))
        # libs.rSleep(500)
        # self.ctl.sendMouse(289, 135)
        # #self.ctl.clickLeft(289, 135)
        # #win32gui.SetForegroundWindow(current_hwnd)
        # win32api.SetCursorPos(current_cursor_pos)

        self.exitToMenu()
        self.waitMenu()
        self.clickContinueGameButton()
        self.waitTown()
        self.goIntoWarrior()
        
        
    def isInTown(self):
        try:
            img = config.imgs_dir + '0_heart.jpg'
            loc = self.ctl.isImageAtRegion(img, 672, 31, 72, 57)
            return True
        except:
            return False
        
    def goIntoWarrior(self):
        self.ctl.move(17, 167)
        libs.sleep(2900)
        self.ctl.select(289, 135)
        libs.sleep(2900)
        
    def exit(self):
        self.ctl.sendESC()
        libs.rSleep(1800)
        self.ctl.clickLeft(397, 281)

    def clickContinueGameButton(self):
        self.ctl.clickLeft(222, 237)

    def needRepair(self):
        color = self.ctl.isColorAtPixel(575, 21, "#fff000")
        if color: return True
        else: return False

    def setRepair(self):
        if self.needRepair():
            config.REPAIRP = True
                
    def repair(self):
        self.ctl.select(491, 64)
        libs.sleep(1400)
        self.ctl.select(422, 120)
        libs.sleep(800)
        self.ctl.select(141, 293)
        libs.sleep(800)
        self.ctl.sendESC()
        libs.sleep(300)
        config.REPAIRP = False

    def isInMenu(self):
        color1 = self.ctl.isColorAtPixel(43, 5, "#050401")
        color2 = self.ctl.isColorAtPixel(749, 589, "#010101")
        if color1 and color2: return True
        else: return False

    def _into_start_point(self):
        if self.isInMenu():
            libs.rSleep(300)
            self.clickContinueGameButton()
            return
        libs.rSleep(300)
        self.exit()
        if self.waitMenu():
            libs.rSleep(300)
            self.clickContinueGameButton()
            return
            
    def waitMenu(self):
        for a in range(60):
            if self.isInMenu(): return True
            libs.rSleep(500)

        return False

    
    def waitTown(self):
        for a in range(60):
            if self.isInTown(): return True
            libs.rSleep(500)

        return False
        

    def exitToMenu(self):
        if self.isInMenu(): return
        self.ctl.sendESC()
        libs.rSleep(500)
        self.ctl.clickLeft(397, 281)
        return self.waitMenu()

    def goFront(self):
        self.ctl.select(757, 123)
        libs.rSleep(2700)
        self.ctl.select(588, 241)
        libs.rSleep(1500)
        self.ctl.move(404, 367)
        libs.rSleep(800)

        self.ctl.move(789, 149)
        libs.rSleep(2400)
        
        self.ctl.move(782, 174)
        libs.rSleep(2400)

        #self.ctl.move(629, 201)
        self.ctl.move(580, 230)
        libs.rSleep(1800)


    def war(self):
        self.ctl.setCursorPos(790, 274)
        libs.sleep(300)
        self.ctl.sendKey("1")
       	libs.rSleep(300)

        self.ctl.setCursorPos(326, 345)
        libs.rSleep(100)
        self.ctl.swingStandingLefta(8, 500)
        self.ctl.setCursorPos(497, 275)
        libs.sleep(100)
        self.ctl.swingStandingRighta(8, 500)

        self.ctl.sendKey("4")
        libs.rSleep(400)
        self.ctl.sendKey("2")        
        libs.sleep(400)

        # Hero
        self.ctl.sendKey("3")        
        libs.rSleep(400)
        self.ctl.sendKey("3")
        libs.rSleep(400)

        # Volcano
        self.ctl.clickMouseRightButton()
        libs.rSleep(400)
        self.ctl.clickMouseRightButton()        

        self.ctl.swingStandingLefta(9, 500)
        self.ctl.swingStandingRighta(9, 500)

        self.ctl.sendKey("2")
        libs.rSleep(200)
        self.ctl.swingStandingLefta(11, 500)
        self.ctl.swingStandingRighta(11, 500)        
        
        #pickItems()
        #imagePickItem()

        self.ctl.move(529, 232)
        libs.rSleep(1200)
        self.ctl.move(40, 374)
        libs.rSleep(1800)
        self.ctl.move(524, 195)
        libs.rSleep(1700)


    def monster(self):
        #(119-120, 246-256, 158-256)
        #contours = self.ctl.findHSVEX([119, 236, 203], [120, 255, 255], 20)
        contours = self.ctl.findHSVEX([119, 166, 116], [120, 255, 255], 20)
        try:
            #print "Length", len(contrours[0])
            for a in range(len(contours)):
                #print "contour length: ", len(contours[a])
                #print contours[a]                
                #print contours[a][0]
                #print contours[a][-1]
                # The result contours[a] is [[[90 393]] [[90 394]] [[91 395]] ...
                x1, y1 = contours[a][0][0]
                x2, y2 = contours[a][-1][0]
                #print "coord", x1, y1, x2, y2
                if 97 < y1 < 470: #and abs(x2 - x1) > 5:
                    print "coordinate: ", x1, y1, x2, y2
                    return x1, y1
            
            # if y > 480 or y < 150:
            #     x, y = contours[0][1][0]
            return

        except:
            return

    def attackMonster(self):
        try:
            x, y = self.monster()
            print "monster", str(x), str(y)
        except:
            return

        self.ctl.setCursorPos(x, y)
        libs.sleep(400)
        for a in range(30):
            self.ctl.skill2()
            libs.sleep(100)

    def war_throw(self):
        self.ctl.skill4()
        libs.sleep(300)
        self.ctl.skill1()
        self.ctl.setCursorPos(782, 227)
        libs.sleep(300)
        for a in range(30):
            # throw
            self.ctl.skill2()
            libs.sleep(100)

        self.ctl.skill3()
        libs.sleep(300)

        self.ctl.skillb()
        libs.sleep(300)

        # for a in range(100):
        #     self.ctl.skill2()
        #     libs.sleep(100)
        # center is about (406, 278)
        # 1step is (491, 273)
        for a in range(3):
            self.ctl.skill1()
            self.attackMonster()
            libs.sleep(700)

        self.ctl.move(556, 278)
        libs.sleep(1000)

        self.ctl.pickItems()

        for a in range(3):
            self.ctl.skill1()
            self.attackMonster()
            libs.sleep(700)

        self.ctl.move(556, 278)
        libs.sleep(1000)

        self.ctl.skilla()
        
        for a in range(3):
            self.ctl.skill1()
            self.attackMonster()
            libs.sleep(700)

        for a in range(3):
            self.ctl.skill1()
            self.attackMonster()
            libs.sleep(700)

        self.ctl.pickItems()


    def setSkill(self):
        if not self.isInMenu():
            self.exitToMenu()
        libs.sleep(800)
        self.clickContinueGameButton()
        self.waitTown()
        libs.sleep(800)
        self.ctl.setSkills(Skills.war_jump)
        libs.sleep(800)
        self.exitToMenu()        
            

def main():
    d3w = DWindowUtiles()
    win32gui.SetForegroundWindow(d3w.hwnd)
    win32gui.MoveWindow(d3w.hwnd, 0, 0, 800, 600, win32con.TRUE)
    controller = D3Controller(D3ForegroundController)
    town = Town(controller)
    libs.sleep(1400)
    town.setSkill()

    while 1:
        town.clickContinueGameButton()
        town.waitTown()
        libs.sleep(800)
        town.ctl.click(776, 211) # remove subject window

        if config.REPAIRP:
            town.repair()
            town.exitToMenu()
            continue
        
        town.goIntoWarrior()
        libs.sleep(1000)
        town.goFront()
        town.setRepair()
            #town.war_throw()
        town.war()

        for a in range(2):
            town.ctl.pickItems()
            libs.sleep(100)

            town.ctl.move(529, 232)
            libs.sleep(1000)
            town.ctl.pickItems()
            
            town.exitToMenu()
    
            
if __name__ == "__main__":
    main()
