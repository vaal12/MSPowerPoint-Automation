#Use 
# C:\Anaconda3\python.exe  C:\Anaconda3\cwp.py C:\Anaconda3 C:\Anaconda3\python.exe c:\Dev\04.Python\06.Office_automation\Powerpoint_presentation\main.py 
#  to run so it loads anaconda dependencies

#Powerpoint reference:
#   https://docs.microsoft.com/en-us/office/vba/api/powerpoint.textrange.text

import time
import win32com.client as win32

import pandas as pd
import numpy as np

import data_requisition

def placePictureOverShape(pic_file_name, PPSlide, PPshape):
    PPSlide.Shapes.AddPicture(
        pic_file_name,
        1, 1, 
        PPshape.Left, PPshape.Top, PPshape.Width, PPshape.Height
    )


import datetime
import os

PP_TEMPLATE_FILE_NAME = r"c:\Dev\04.Python\06.Office_automation\Powerpoint_presentation\Template presentation"
def pp():
    import win32com.client, sys
    Application = win32com.client.Dispatch("PowerPoint.Application")
    Application.Visible = True
    Presentation = Application.Presentations.Open(PP_TEMPLATE_FILE_NAME+".pptx")

    curr_date = datetime.datetime.now().strftime("%d-%m-%Y")
    Presentation.Slides(1).Shapes[1].TextFrame.TextRange.Text= curr_date

    Presentation.Slides(2).Shapes[0].TextFrame.TextRange.Text= "Total company sales by month"

    for Slide in Presentation.Slides:
        print("Tranversing slide")
        for Shape in Slide.Shapes:
            print("tranversing shape:{}".format(Shape.Type))
            #MsoShapeType: https://docs.microsoft.com/en-us/office/vba/api/office.msoshapetype
            # print("Have text range:{}".format(Shape.TextFrame.TextRange))
            txt_str = str(Shape.TextFrame.TextRange.Text)
            print("txt_str type:{}".format(txt_str))
            if txt_str.find("%%Placeholder%%")>-1:
                # print("FOUND")
                # print("Top:{}    Left:{}".format(Shape.Top, Shape.Left))
                # print("Width:{}   Height:{}".format(Shape.Width, Shape.Height))
                # Typical figures:
                #     Top:126.0    Left:36.0
                #     Width:648.0   Height:356.3750305175781

                # pic = data_requisition.getIrisAnalysisPlot()
                pic = data_requisition.getMonthlySales()
                placePictureOverShape(
                    pic,
                    Slide, Shape
                )# THis will add picture to the Shapes immediately, so the loop will not stop - have to break after this
                os.remove(pic)
                break
            # Shape.TextFrame.TextRange.Font.Name = "Arial"
    
    pptLayout = Presentation.Slides(2).CustomLayout
    newSlide = Presentation.Slides.AddSlide(Presentation.Slides.Count+1, pptLayout)
    newSlide.Shapes[0].TextFrame.TextRange.Text = "Sales of the top (by sales) 20 artists "
    pic = data_requisition.getTop20ArtistSales()
    placePictureOverShape(
        pic,
        newSlide, newSlide.Shapes[1]
    )
    os.remove(pic)

    newSlide = Presentation.Slides.AddSlide(Presentation.Slides.Count+1, pptLayout)
    newSlide.Shapes[0].TextFrame.TextRange.Text = "Proportion of sales by artists"
    pic = data_requisition.getDistributionOfSales()
    placePictureOverShape(
        pic,
        newSlide, newSlide.Shapes[1]
    )
    os.remove(pic)

    for employee in data_requisition.getEmployeesList():
        newSlide = Presentation.Slides.AddSlide(Presentation.Slides.Count+1, pptLayout)
        newSlide.Shapes[0].TextFrame.TextRange.Text = "Employee details:{} {}".format(
            employee["FirstName"], employee["LastName"])
        pic = data_requisition.getEmployeeSalesGraph(employee["EmployeeId"])
        placePictureOverShape(
            pic,
            newSlide, newSlide.Shapes[1]
        )
        os.remove(pic)



    

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
# def createFigure():
#     NUM_OF_SAMPLES = 10
#     x_arr = np.random.rand(NUM_OF_SAMPLES)
#     y_arr = np.random.rand(NUM_OF_SAMPLES)
#     idx = range(NUM_OF_SAMPLES)

#     df = pd.DataFrame(list(zip(x_arr, y_arr)), index=idx, columns=["x", "y"])
#     ax = df.plot.scatter(x="x", y="y", figsize=(12,6))
#     fig = ax.get_figure()
#     fig.savefig(r"c:\Dev\04.Python\06.Office_automation\Powerpoint_presentation\fig1.png")
    
    
 
if __name__ == "__main__":
    pass
    # excel()
    # createFigure()
    pp()