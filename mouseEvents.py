import time

from Quartz.CoreGraphics import CGEventCreateMouseEvent
from Quartz.CoreGraphics import CGEventPost
from Quartz.CoreGraphics import kCGEventMouseMoved
from Quartz.CoreGraphics import kCGEventLeftMouseDown
from Quartz.CoreGraphics import kCGEventLeftMouseUp
from Quartz.CoreGraphics import kCGMouseButtonLeft
from Quartz.CoreGraphics import kCGHIDEventTap

def mouseEvent(type, posx, posy):
	theEvent = CGEventCreateMouseEvent(
		None, 
		type, 
		(posx,posy), 
		kCGMouseButtonLeft)
	CGEventPost(kCGHIDEventTap, theEvent)

def mouseMove(posx,posy):
	mouseEvent(kCGEventMouseMoved, posx,posy)

def mouseClick(posx,posy):
	mouseEvent(kCGEventLeftMouseDown, posx,posy)
	mouseEvent(kCGEventLeftMouseUp, posx,posy)

# def mouseDrag(x1,y1,x2,y2):
# 	mouseEvent(kCGEventLeftMouseDown, x1,y1)
# 	mouseEvent(kCGEventLeftMouseUp, x2,y2)
def mouseDown(posx,posy):
	mouseEvent(kCGEventLeftMouseDown, posx,posy)

def mouseUp(posx,posy):
	mouseEvent(kCGEventLeftMouseUp, posx,posy)

def mouseDrag(x1,y1,x2,y2):
	mouseClick(x1,y1)
	mouseEvent(kCGEventLeftMouseDown,x1,y1)
	time.sleep(.25)
	mouseMove(x2,y2)
	time.sleep(.25)
 	mouseEvent(kCGEventLeftMouseUp,x2,y2)
