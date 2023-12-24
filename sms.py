import pandas as pd
import pywhatkit
import pyautogui
import time

excel_file_path = 'E:\\W_A_M-1\\Number.xlsx'

try:
    df = pd.read_excel(excel_file_path)

    for index, row in df.iterrows():
        mobile_number = "+880" + str(row['Number'])
        image_path = "E:\\W_A_M-1\\Plastine.jpeg"
        caption = "Prove that 'Humanity is still alive'. Please donate to Palestine."

        print(f"Sending image to {mobile_number}")
        pywhatkit.sendwhats_image(mobile_number, image_path, caption, 20)

        time.sleep(10)  # Adjust the delay based on your internet speed and processing time

        # Move the mouse to the coordinates of the close button and click
        pyautogui.click(x=1917, y=10)  # Adjust the coordinates based on your screen resolution

        time.sleep(2)

except Exception as e:
    print(f"An error occurred: {e}")