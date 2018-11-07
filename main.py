#Use 
# C:\Anaconda3\python.exe  C:\Anaconda3\cwp.py C:\Anaconda3 C:\Anaconda3\python.exe main.py
#  to run so it loads anaconda dependencies

import time
import win32com.client as win32

import pandas as pd
import numpy as np

def placePictureOverShape(pic_file_name, PPSlide, PPshape):
    PPSlide.Shapes.AddPicture(
        pic_file_name,
        1, 1, 
        PPshape.Left, PPshape.Top, PPshape.Width, PPshape.Height
    )


import datetime

PP_TEMPLATE_FILE_NAME = r"c:\Dev\04.Python\06.Office_automation\Powerpoint_presentation\Template presentation_2"
def pp():
    import win32com.client, sys
    Application = win32com.client.Dispatch("PowerPoint.Application")
    Application.Visible = True
    Presentation = Application.Presentations.Open(PP_TEMPLATE_FILE_NAME+".pptx")

    curr_date = datetime.datetime.now().strftime("%d-%m-%Y")
    Presentation.Slides(1).Shapes[1].TextFrame.TextRange.Text= curr_date

    for Slide in Presentation.Slides:
        for Shape in Slide.Shapes:
                print("Have text range:{}".format(Shape.TextFrame.TextRange))
                txt_str = str(Shape.TextFrame.TextRange)
                print("txt_str type:{}".format(type(txt_str)))
                if txt_str.find("%%Placeholder%%")>-1:
                    print("FOUND")
                    print("Top:{}    Left:{}".format(Shape.Top, Shape.Left))
                    print("Width:{}   Height:{}".format(Shape.Width, Shape.Height))
                    # Typical figures:
                    #     Top:126.0    Left:36.0
                    #     Width:648.0   Height:356.3750305175781
                    placePictureOverShape(
                        r"c:\Dev\04.Python\06.Office_automation\Powerpoint_presentation\iris.png",
                        Slide, Shape
                    )
                    # Slide.Shapes.AddPicture(r"c:\Dev\04.Python\06.Office_automation\Powerpoint_presentation\iris.png",
                    #     1, 1, Shape.Left, Shape.Top, Shape.Width, Shape.Height)
                    break
                # Shape.TextFrame.TextRange.Font.Name = "Arial"
    
    pptLayout = Presentation.Slides(2).CustomLayout
    newSlide = Presentation.Slides.AddSlide(Presentation.Slides.Count+1, pptLayout)
    newSlide.Shapes[0].TextFrame.TextRange.Text = "New slide 1"
    placePictureOverShape(
        r"c:\Dev\04.Python\06.Office_automation\Powerpoint_presentation\top20sales.png",
        newSlide, newSlide.Shapes[1]
    )
    # 

    newSlide = Presentation.Slides.AddSlide(Presentation.Slides.Count+1, pptLayout)
    newSlide.Shapes[0].TextFrame.TextRange.Text = "New slide 2"
    placePictureOverShape(
        r"c:\Dev\04.Python\06.Office_automation\Powerpoint_presentation\distribution_sales.png",
        newSlide, newSlide.Shapes[1]
    )

    

    Presentation.SaveAs(PP_TEMPLATE_FILE_NAME+"_"+curr_date+".pptx")
    #https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppsaveasfiletype
    Presentation.SaveAs(PP_TEMPLATE_FILE_NAME+"_"+curr_date+".pdf", 32)
    # Application.Quit()

def excel():
    xl = win32.gencache.EnsureDispatch('PowerPoint.Application')
    xl.Show()
    ss = xl.Workbooks.Add()
    sh = ss.ActiveSheet
 
    xl.Visible = True
    time.sleep(1)
 
    sh.Cells(1,1).Value = 'Hacking Excel with Python Demo'
 
    time.sleep(1)
    for i in range(2,8):
        sh.Cells(i,1).Value = 'Line %i' % i
        time.sleep(1)
 
    ss.Close(False)
    xl.Application.Quit()

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
def createFigure():
    NUM_OF_SAMPLES = 10
    x_arr = np.random.rand(NUM_OF_SAMPLES)
    y_arr = np.random.rand(NUM_OF_SAMPLES)
    idx = range(NUM_OF_SAMPLES)

    df = pd.DataFrame(list(zip(x_arr, y_arr)), index=idx, columns=["x", "y"])
    ax = df.plot.scatter(x="x", y="y", figsize=(12,6))
    fig = ax.get_figure()
    fig.savefig(r"c:\Dev\04.Python\06.Office_automation\Powerpoint_presentation\fig1.png")
    
    
 
if __name__ == "__main__":
    pass
    # excel()
    createFigure()
    pp()