#!/usr/bin/python
# coding: utf-8

import os
import time
import struct

import win32gui
import win32api
import win32con
import win32ui
#from ctypes import *
import ctypes
import cv2
from cv2 import cv
import numpy

import libs

class WindowSingleton:
    __shared_state = {}
    def __init__(self):
        self.__dict__ = self.__shared_state

    def __len__(self):
        return len(self.__shared_state)

    def __getitem__(self, key):
        return self.__shared_state[key]

class DWindow(object):
    '''
    >>> dw = DWindow(u'Notepd')
    >>> dw.hwnd #doctest: +SKIP
    >>> win32gui.MoveWindow(dw.hwnd, 2100, 0, 800, 600, win32con.TRUE)
    >>> dw.rects
    ((2100, 0, 2908, 627), (-2104, -23), (800, 600), (4, 23))
    >>> dw.left
    2104
    >>> dw.top
    23
    >>> dw.dx
    4
    >>> dw.dy
    23
    >>> dir(dw.window)
    ['_WindowSingleton__shared_state', '__doc__', '__getitem__', '__init__', '__len__', '__module__', 'dc', 'dx', 'dy', 'height', 'hwnd', 'left', 'rects', 'top', 'width', 'window_left', 'window_top']
    '''

    def __init__(self, process_name):
        self.process_name = process_name
        self._init()

    def _init(self):
        self.window = WindowSingleton()
        self._setHwnd()
        self._setDc()
        self._setRects()
        if len(self.window): return


    def __getattr__(self, key):
        return self.window[key]

    
    def _setHwnd(self):
        # TODO: exception
        self.window.hwnd = win32gui.FindWindow(None, self.process_name)

    def _setDc(self):
        self.window.dc = win32gui.GetWindowDC(self.hwnd)

    def _setRects(self):
        window_coord = win32gui.GetWindowRect(self.hwnd)
        # There are ClientToScreen, MapWindowPoints
        client_coord = win32gui.ScreenToClient(self.hwnd, (0, 0))

        # These are the coordinates for client
        # The client coordinates that exclude title bar
        _, _, self.window.width, self.window.height = win32gui.GetClientRect(self.hwnd)
        self.window.left, self.window.top = [abs(a) for a in client_coord]

        # The window coordinates
        self.window.window_left = window_coord[0]
        self.window.window_top = window_coord[1]

        # Height of title bar and width of left boarder
        self.window.dx = abs(self.left - self.window_left)
        self.window.dy = abs(self.top - self.window_top)

        self.window.rects = (window_coord, client_coord, (self.width, self.height), (self.dx, self.dy))


class DWindowUtilesBase(DWindow):
    '''
     - We deals client's coordinate.

    ### === getPixel
    ### __________________________________________________________
    >>> dwu = DWindowUtilesBase(u'Notepd')
    >>> dx = 100
    >>> dy = 100
    >>> our_color = dwu.getPixel(dx, dy)

    >>> x = dwu.left + dx
    >>> y = dwu.top + dy
    >>> dhwnd = win32gui.GetDesktopWindow()
    >>> dhdc = win32gui.GetWindowDC(dhwnd)
    >>> expected_color = dwu.getPixelWithDC(dhdc, x, y)

    >>> our_color = expected_color

    >>> dwu.saveScreenshot("test") #doctest: +SKIP

    >>> dwu.sendKey("t") #doctest: +SKIP

    >>> dwu.clickLeft(799, 599) #doctest: +SKIP

    >>> dwu.setCursorPos(700, 300) #doctest: +SKIP
    '''

    def getPixel(self, x, y):
        return self.getPixelWithDC(self.dc, x, y)

    def getPixelWithDC(self, dc, x, y):
        # GetPixel include title bar
        x += self.dx
        y += self.dy
        long_color = win32gui.GetPixel(dc, x, y)
        color = int(long_color)
        color = (color & 0xff), ((color >> 8) & 0xff), ((color >> 16) & 0xff)
        return libs.rgb_to_hex(color)

    def getPixelOnDesktop(self, x, y):
        hwnd = win32gui.GetDesktopWindow()
        dc = win32gui.GetWindowDC(self.desktopHwnd)
        long_color = win32gui.GetPixel(dc, x, y)
        color = int(long_color)
        color = (color & 0xff), ((color >> 8) & 0xff), ((color >> 16) & 0xff)
        return libs.rgb_to_hex(color)


    def findPixel(self, color):
        pass

    def _findPixelType1(self, color):
        hdc = win32gui.GetWindowDC(self.hwnd)
        dc_obj = win32ui.CreateDCFromHandle(hdc)
        memorydc = dc_obj.CreateCompatibleDC()

        data_bitmap = win32ui.CreateBitmap()
        data_bitmap.CreateCompatibleBitmap(dc_obj, self.width, self.height)

        memorydc.SelectObject(data_bitmap)
        memorydc.BitBlt((0, 0), (self.width, self.height), dc_obj, (self.dx, self.dy), win32con.SRCCOPY)

        bmpheader = struct.pack("LHHHH", struct.calcsize("LHHHH"),
                                self.width, self.height, 1, 24)
        c_bmpheader = ctypes.create_string_buffer(bmpheader)

        # padded_length = (string_length + 3) & -3 for 4-byte aligned.
        c_bits = ctypes.create_string_buffer(" " * (self.width * ((self.height * 3 + 3) & -3)))

        res = ctypes.windll.gdi32.GetDIBits(memorydc.GetSafeHdc(), data_bitmap.GetHandle(),
                                            0, self.height, c_bits, c_bmpheader, win32con.DIB_RGB_COLORS)

        win32gui.DeleteDC(hdc)
        win32gui.ReleaseDC(self.hwnd, hdc)
        memorydc.DeleteDC()
        win32gui.DeleteObject(data_bitmap.GetHandle())

        cv_im = cv.CreateImageHeader((self.width, self.height), cv.IPL_DEPTH_8U, 3)
        cv.SetData(cv_im, c_bits.raw)
        # flip around x-axis
        cv.Flip(cv_im, None, 0)

        return repr(cv_im[1][1])
        

    def saveScreenshot(self, name):
        dcObj = win32ui.CreateDCFromHandle(self.dc)
        cdc = dcObj.CreateCompatibleDC()
        dataBitMap = win32ui.CreateBitmap()

        dataBitMap.CreateCompatibleBitmap(dcObj, self.width, self.height)
        cdc.SelectObject(dataBitMap)
        cdc.BitBlt((0, 0),(self.width, self.height) , dcObj, (self.dx, self.dy), win32con.SRCCOPY)
        image_name = os.getcwd() + "\\" + name + str(int(time.time())) + '.bmp'
        dataBitMap.SaveBitmapFile(cdc, image_name)
        cdc.DeleteDC()

    def sendKey(self, key):
        try:
            key = ord(key.upper())
        except AttributeError:
            pass

        scan_code = ctypes.windll.user32.MapVirtualKeyA(key, 0)

        lparam_down = 0x00000001 | (scan_code << 16)
        lparam_up = 0xC0000001 | (scan_code << 16)

        win32api.PostMessage(self.hwnd, win32con.WM_KEYDOWN, key, lparam_down)
        win32api.PostMessage(self.hwnd, win32con.WM_KEYUP, key, lparam_up)

    def sendMouseLeftClick(self, x, y):
        lparam = win32api.MAKELONG(x, y)

        win32api.PostMessage(self.hwnd, win32con.WM_LBUTTONDOWN, win32con.VK_LBUTTON, lparam)
        win32api.PostMessage(self.hwnd, win32con.WM_LBUTTONUP, win32con.VK_LBUTTON, lparam)

    def sendMouseLeftClickEx(self, x, y, count=1, shift=False, sleep=300):
        lparam = win32api.MAKELONG(x, y)

        if not shift:
            win32api.PostMessage(self.hwnd, win32con.WM_LBUTTONDOWN, win32con.VK_LBUTTON, lparam)
            win32api.PostMessage(self.hwnd, win32con.WM_LBUTTONUP, win32con.VK_LBUTTON, lparam)
            return

        win32api.PostMessage(self.hwnd, win32con.WM_KEYDOWN, win32con.VK_SHIFT, 0)

        for a in range(count):
            win32api.PostMessage(self.hwnd, win32con.WM_LBUTTONDOWN, win32con.VK_LBUTTON, lparam)
            win32api.PostMessage(self.hwnd, win32con.WM_LBUTTONUP, win32con.VK_LBUTTON, lparam)
            libs.rSleep(300)

        win32api.PostMessage(self.hwnd, win32con.WM_KEYUP, win32con.VK_SHIFT, 0)
        
    def sendMouseRightClick(self, x, y):
        lparam = win32api.MAKELONG(x, y)

        win32api.PostMessage(self.hwnd, win32con.WM_RBUTTONDOWN, win32con.VK_RBUTTON, lparam)
        win32api.PostMessage(self.hwnd, win32con.WM_RBUTTONUP, win32con.VK_RBUTTON, lparam)


    def sendMouseRightClickEx(self, x=300, y=300, count=1, shift=False, sleep=300):
        lparam = win32api.MAKELONG(x, y)

        if not shift:
            win32api.PostMessage(self.hwnd, win32con.WM_RBUTTONDOWN, win32con.VK_RBUTTON, lparam)
            win32api.PostMessage(self.hwnd, win32con.WM_RBUTTONUP, win32con.VK_RBUTTON, lparam)
            return

        win32api.PostMessage(self.hwnd, win32con.WM_KEYDOWN, win32con.VK_SHIFT, 0)

        for a in range(count):
            win32api.PostMessage(self.hwnd, win32con.WM_RBUTTONDOWN, win32con.VK_RBUTTON, lparam)
            win32api.PostMessage(self.hwnd, win32con.WM_RBUTTONUP, win32con.VK_RBUTTON, lparam)
            libs.rSleep(300)

        win32api.PostMessage(self.hwnd, win32con.WM_KEYUP, win32con.VK_SHIFT, 0)

        
        if shift:
            wparam = win32con.MK_SHIFT | win32con.MK_RBUTTON
        else:
            wparam = win32con.MK_RBUTTON

        for a in range(count):
            win32api.PostMessage(self.hwnd, win32con.WM_RBUTTONDOWN, wparam, lparam)
            win32api.PostMessage(self.hwnd, win32con.WM_RBUTTONUP, wparam, lparam)
            libs.sleep(sleep)
        
        

    def sendMouseMiddleClick(self, x, y):
        lparam = win32api.MAKELONG(x, y)

        win32api.PostMessage(self.hwnd, win32con.WM_MBUTTONDOWN, win32con.VK_MBUTTON, lparam)
        win32api.PostMessage(self.hwnd, win32con.WM_MBUTTONUP, win32con.VK_MBUTTON, lparam)


    # Not used
    def __sendMouseLeftClick(self, x=300, y=300, count=1, shift=False):
        lparam = win32api.MAKELONG(x, y)

        if shift:
            wparam = win32con.MK_SHIFT | win32con.MK_LBUTTON
        else:
            wparam = win32con.MK_LBUTTON

        for a in range(count):
            win32api.PostMessage(self.hwnd, win32con.WM_LBUTTONDOWN, wparam, lparam)
            win32api.PostMessage(self.hwnd, win32con.WM_LBUTTONUP, wparam, lparam)
            libs.rSleep(300)
        
        # if not shift:
        #     win32api.PostMessage(self.hwnd, win32con.WM_LBUTTONDOWN, win32con.VK_LBUTTON, lparam)
        #     win32api.PostMessage(self.hwnd, win32con.WM_LBUTTONUP, win32con.VK_LBUTTON, lparam)
        #     return

        # win32api.PostMessage(self.hwnd, win32con.WM_KEYDOWN, win32con.VK_SHIFT, 0)
        # for a in range(count):
        #     win32api.PostMessage(self.hwnd, win32con.WM_LBUTTONDOWN, win32con.VK_LBUTTON, lparam)
        #     win32api.PostMessage(self.hwnd, win32con.WM_LBUTTONUP, win32con.VK_LBUTTON, lparam)
        #     libs.rSleep(300)

        # win32api.PostMessage(self.hwnd, win32con.WM_KEYUP, win32con.VK_SHIFT, 0)

    # Not used
    def sendMouseLeftClick2(self):
        #(388, 236)

        class POINT(Structure):
            _fields_ = [("x", c_ulong),
                        ("y", c_ulong)]

        #byref(pt)
        pt = POINT()
        pt.x = 126 #396
        pt.y = 225 #233

        pt2 = POINT()
        pt.x = 0
        pt.y = 0
        
        #lparam = win32api.MAKELONG(388, 236)
        lparam = win32api.MAKELONG(0, 0)
        lparam2 = win32api.MAKELONG(126, 225)
        
        libs.rSleep(300)
        #win32api.PostMessage(self.hwnd, win32con.WM_NCLBUTTONDOWN, win32con.HTCLIENT, pt)
        #win32api.PostMessage(self.hwnd, win32con.WM_NCLBUTTONUP, win32con.HTCLIENT, pt)
        # windll.user32.PostMessageA(self.hwnd, win32con.WM_NCLBUTTONDOWN, win32con.VK_LBUTTON, pt)
        # windll.user32.PostMessageA(self.hwnd, win32con.WM_NCLBUTTONUP, win32con.VK_LBUTTON, pt)
        # win32api.PostMessage(self.hwnd, win32con.WM_NCMOUSEMOVE, win32con.HTCLIENT, pt)
        # win32api.PostMessage(self.hwnd, win32con.WM_MOUSEMOVE, win32con.MK_LBUTTON, lparam2)
        # win32api.PostMessage(self.hwnd, win32con.WM_LBUTTONDOWN, win32con.VK_LBUTTON, lparam2)
        # win32api.PostMessage(self.hwnd, win32con.WM_LBUTTONUP, win32con.VK_LBUTTON, lparam2)
        # win32api.PostMessage(self.hwnd, win32con.WM_MOUSEO, win32con.VK_LBUTTON, lparam2)
        # win32api.PostMessage(self.hwnd, win32con.WM_LBUTTONUP, win32con.VK_LBUTTON, lparam2)

        x, y = self._getDesktopCoordinate(126, 225)
        self.setCursorPos(396, 233)
        libs.rSleep(200)
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0)
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0)


        
                
        
    def clickMouse(self, mouse_down_event, mouse_up_event, x, y):

        x += self.left
        y += self.top

        current_hwnd = win32gui.GetForegroundWindow()
        current_cursor_pos = win32gui.GetCursorPos()

        #win32gui.SetForegroundWindow(self.hwnd)
        win32api.ClipCursor((x-1,y-1,x+1,y+1))
        win32api.SetCursorPos((x,y))
        win32api.mouse_event(mouse_down_event| \
                                win32con.MOUSEEVENTF_ABSOLUTE,0,0)
        win32api.mouse_event(mouse_up_event| \
                                win32con.MOUSEEVENTF_ABSOLUTE,0,0)
        win32api.ClipCursor((0,0,0,0))

        win32gui.SetForegroundWindow(current_hwnd)
        win32api.SetCursorPos(current_cursor_pos)

    def _getDesktopCoordinate(self, x, y):
        x += self.left
        y += self.top
        return x, y

    def posToDesktopPos(self, x, y):
        return self._getDesktopCoordinate(x, y)

    def desktopPosToPos(self, x, y):
        x -= self.left
        y -= self.top
        return x, y

    def setCursorPos(self, x, y):
        # FIXME: There is no action with doctest and time.sleep
        pos = self.posToDesktopPos(x, y)
        win32api.SetCursorPos(pos)

    def setCursorAbsolutePos(self, x, y):
        win32api.SetCursorPos((x, y))

    def getCurrentCursorPos(self):
        x, y = win32gui.GetCursorPos()
        dw_right_top_x = self.left + self.width
        dw_left_bottom_y = self.top + self.height

        if (self.left < abs(x) < dw_right_top_x) and (self.top < abs(y) < dw_left_bottom_y):
            # Cursor on window
            return self.desktopPosToPos(x, y)
        # Out of window
        return None

    def getPixelCurrentPos(self):
        'Work on window. Other will return None.'
        try:
            x, y = self.getCurrentCursorPos()
            return self.getPixel(x, y)
        except TypeError:
            return None

    def getPixelInfo(self):
        "Work on Desktop."
        pos = self.getCurrentCursorPos()
        if not pos:
            pos = win32gui.GetCursorPos()
            color = self.getPixelOnDesktop(*pos)
            return (pos, color)
        color = self.getPixel(*pos)
        return (pos, color)
            
    def findImage(self, source):
        hdc = win32gui.GetWindowDC(self.hwnd)
        dc_obj = win32ui.CreateDCFromHandle(hdc)
        memorydc = dc_obj.CreateCompatibleDC()
        data_bitmap = win32ui.CreateBitmap()
        data_bitmap.CreateCompatibleBitmap(dc_obj, self.width, self.height)

        memorydc.SelectObject(data_bitmap)
        memorydc.BitBlt((0, 0), (self.width, self.height), dc_obj, (self.dx, self.dy), win32con.SRCCOPY)

        bmpheader = struct.pack("LHHHH", struct.calcsize("LHHHH"),
                                self.width, self.height, 1, 24)
        c_bmpheader = ctypes.create_string_buffer(bmpheader)

        # padded_length = (string_length + 3) & -3 for 4-byte aligned.
        c_bits = ctypes.create_string_buffer(" " * (self.width * ((self.height * 3 + 3) & -3)))

        res = ctypes.windll.gdi32.GetDIBits(memorydc.GetSafeHdc(), data_bitmap.GetHandle(),
                                            0, self.height, c_bits, c_bmpheader, win32con.DIB_RGB_COLORS)

        win32gui.DeleteDC(hdc)
        win32gui.ReleaseDC(self.hwnd, hdc)
        memorydc.DeleteDC()
        win32gui.DeleteObject(data_bitmap.GetHandle())

        cv_im = cv.CreateImageHeader((self.width, self.height), cv.IPL_DEPTH_8U, 3)
        cv.SetData(cv_im, c_bits.raw)
        # flip around x-axis
        cv.Flip(cv_im, None, 0)

        template_source = cv.LoadImage(source)

        # From the manual of MatchTemplate
        result_width = cv_im.width - template_source.width + 1
        result_height = cv_im.height - template_source.height + 1;
        result = cv.CreateImage((result_width, result_height), 32, 1)

        cv.MatchTemplate(cv_im, template_source, result, cv2.TM_CCOEFF_NORMED)
        return cv.MinMaxLoc(result)

    def findImageEx(self, source, x, y, width, height):
        hdc = win32gui.GetWindowDC(self.hwnd)
        dc_obj = win32ui.CreateDCFromHandle(hdc)
        memorydc = dc_obj.CreateCompatibleDC()
        data_bitmap = win32ui.CreateBitmap()
        data_bitmap.CreateCompatibleBitmap(dc_obj, self.width, self.height)

        memorydc.SelectObject(data_bitmap)
        memorydc.BitBlt((0, 0), (self.width, self.height), dc_obj, (self.dx, self.dy), win32con.SRCCOPY)

        bmpheader = struct.pack("LHHHH", struct.calcsize("LHHHH"),
                                self.width, self.height, 1, 24)
        c_bmpheader = ctypes.create_string_buffer(bmpheader)

        # padded_length = (string_length + 3) & -3 for 4-byte aligned.
        c_bits = ctypes.create_string_buffer(" " * (self.width * ((self.height * 3 + 3) & -3)))

        res = ctypes.windll.gdi32.GetDIBits(memorydc.GetSafeHdc(), data_bitmap.GetHandle(),
                                            0, self.height, c_bits, c_bmpheader, win32con.DIB_RGB_COLORS)

        win32gui.DeleteDC(hdc)
        win32gui.ReleaseDC(self.hwnd, hdc)
        memorydc.DeleteDC()
        win32gui.DeleteObject(data_bitmap.GetHandle())

        cv_im = cv.CreateImageHeader((self.width, self.height), cv.IPL_DEPTH_8U, 3)
        cv.SetData(cv_im, c_bits.raw)
        # flip around x-axis
        cv.Flip(cv_im, None, 0)

        im_region = cv.GetSubRect(cv_im, (x, y, width, height))
        #cv.SaveImage('aaak.bmp', im_region)

        template_source = cv.LoadImage(source)

        # From the manual of MatchTemplate
        result_width = im_region.width - template_source.width + 1
        result_height = im_region.height - template_source.height + 1;
        result = cv.CreateImage((result_width, result_height), 32, 1)

        cv.MatchTemplate(im_region, template_source, result, cv2.TM_CCOEFF_NORMED)
        minVal, maxVal, minLoc, maxLoc = cv.MinMaxLoc(result)
        #print minVal, maxVal, minLoc, maxLoc

        minLoc2 = minLoc[0] + x, minLoc[1] + y
        maxLoc2 = maxLoc[0] + x, maxLoc[1] + y

        return minVal, maxVal, minLoc2, maxLoc2
        
        
    

    def getMat(self):
        hdc = win32gui.GetWindowDC(self.hwnd)
        dc_obj = win32ui.CreateDCFromHandle(hdc)
        memorydc = dc_obj.CreateCompatibleDC()

        data_bitmap = win32ui.CreateBitmap()
        data_bitmap.CreateCompatibleBitmap(dc_obj, self.width, self.height)

        memorydc.SelectObject(data_bitmap)
        memorydc.BitBlt((0, 0), (self.width, self.height), dc_obj, (self.dx, self.dy), win32con.SRCCOPY)

        bmpheader = struct.pack("LHHHH", struct.calcsize("LHHHH"),
                                self.width, self.height, 1, 24)
        c_bmpheader = ctypes.create_string_buffer(bmpheader)

        # padded_length = (string_length + 3) & -3 for 4-byte aligned.
        c_bits = ctypes.create_string_buffer(" " * (self.width * ((self.height * 3 + 3) & -3)))

        res = ctypes.windll.gdi32.GetDIBits(memorydc.GetSafeHdc(), data_bitmap.GetHandle(),
                                            0, self.height, c_bits, c_bmpheader, win32con.DIB_RGB_COLORS)

        win32gui.DeleteDC(hdc)
        win32gui.ReleaseDC(self.hwnd, hdc)
        memorydc.DeleteDC()
        win32gui.DeleteObject(data_bitmap.GetHandle())

        cv_im = cv.CreateImageHeader((self.width, self.height), cv.IPL_DEPTH_8U, 3)
        cv.SetData(cv_im, c_bits.raw)
        # flip around x-axis
        cv.Flip(cv_im, None, 0)

        mat = cv.GetMat(cv_im)
        return numpy.asarray(mat)

    def findHSV(self, min_hsv, max_hsv):
        result = []
        img = self.getMat()
        img = cv2.cvtColor(img, cv2.COLOR_RGB2HSV)
        lower = numpy.asarray(min_hsv)
        upper = numpy.asarray(max_hsv)
        mask = cv2.inRange(img, lower, upper)
        contours, _ = cv2.findContours(mask, cv2.RETR_LIST, cv2.CHAIN_APPROX_SIMPLE)
        for contour in contours:
            if len(contour) > 3:
                result.append(contour)
        return result

    
    def findHSVEX(self, min_hsv, max_hsv, min_pixels_on_contour):
        result = []
        img = self.getMat()
        img = cv2.cvtColor(img, cv2.COLOR_RGB2HSV)
        lower = numpy.asarray(min_hsv)
        upper = numpy.asarray(max_hsv)
        mask = cv2.inRange(img, lower, upper)
        contours, _ = cv2.findContours(mask, cv2.RETR_LIST, cv2.CHAIN_APPROX_SIMPLE)
        #print 'contours', contours
        for contour in contours:
            if len(contour) > min_pixels_on_contour:
                result.append(contour)
        return result

    def findHSVEX2(self, min_hsv, max_hsv, min_pixels_on_contour):
        result = []
        img = self.getMat()
        img = cv2.cvtColor(img, cv2.COLOR_RGB2HSV)
        lower = numpy.asarray(min_hsv)
        upper = numpy.asarray(max_hsv)
        mask = cv2.inRange(img, lower, upper)
        contours, _ = cv2.findContours(mask, cv2.RETR_LIST, cv2.CHAIN_APPROX_SIMPLE)
        for contour in contours:
            if len(contour) > min_pixels_on_contour:
                result.append(contour)
        return result

    
    
class DWindowUtiles(DWindowUtilesBase):

    def isColorAtPixel(self, x, y, color):
        color = color.lower()
        pos_color = self.getPixel(x, y)
        if color == pos_color: return True
        else: return False

    def isImageAtWindow(self, img):
        minVal, maxVal, minLoc, maxLoc = self.findImage(img)

        if maxVal > 0.75:
            return maxLoc
        raise Exception

    def isImageAtRegion(self, img, x, y, width, height):
        minVal, maxVal, minLoc, maxLoc = self.findImageEx(img, x, y, width, height)
        #print maxVal, minVal
        if maxVal > 0.75:
            return maxLoc
        raise Exception

    def sendESC(self):
        self.sendKey(win32con.VK_ESCAPE)

    def sendAlt(self):
        self.sendKey(win32con.VK_MENU)


    def select(self, x, y):
        self.setCursorPos(x,y)
        libs.rSleep(300)
        self.sendMouse(x, y)



    def clickLeft(self, x, y):
        self.clickMouse(win32con.MOUSEEVENTF_LEFTDOWN, win32con.MOUSEEVENTF_LEFTUP, x, y)

    def clickRight(self, x, y):
        self.clickMouse(win32con.MOUSEEVENTF_RIGHTDOWN, win32con.MOUSEEVENTF_RIGHTUP, x, y)

    def clickMiddle(self, x, y):
        self.clickMouse(win32con.MOUSEEVENTF_MIDDLEDOWN, win32con.MOUSEEVENTF_MIDDLEUP, x, y)


    def clickMouseLeftButton(self):
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0)
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0)

    def clickMouseRightButton(self):
        win32api.mouse_event(win32con.MOUSEEVENTF_RIGHTDOWN, 0, 0)
        win32api.mouse_event(win32con.MOUSEEVENTF_RIGHTUP, 0, 0)

    def clickMouseMiddleButton(self):
        win32api.mouse_event(win32con.MOUSEEVENTF_MIDDLEDOWN, 0, 0)
        win32api.mouse_event(win32con.MOUSEEVENTF_MIDDLEUP, 0, 0)
