from engine import Scraper
import wx
from gui import MainWindow
import logging


if __name__ == "__main__":
    logging.basicConfig(filename='crawler.log', level=logging.INFO, filemode='a', format='%(asctime)s - %(levelname)s - %(message)s')
    logging.info(f"Starting app...")
    eng = Scraper()
    version = "1.0.2"
    default_title = f"Dissolved/Active - Georgia State v{version}"
    app = wx.App(False)
    frame = MainWindow(None, default_title, eng)
    eng.frame = frame
    frame.Show(True) 
    app.MainLoop()
