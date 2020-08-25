# AutoNC-pyautogui
using pyautogui module to fulfill the NC system automatically


highlight:
using two functions to locate a specific image with a certain confidence.
		pyautogui.screenshot()
		(x,y)=pyautogui.center(pyautogui.locateOnScreen(target_image,confidence=0.8))
		pyautogui.moveTo(x, y, 0.5, pyautogui.easeOutQuad) 
		pyautogui.move(x_move, y_move)
		pyautogui.click()