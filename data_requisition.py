import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

PIC_BASE_PATH = "c:\\Dev\\04.Python\\06.Office_automation\\Powerpoint_presentation\\"

def getRandomScatterPlot():
    NUM_OF_SAMPLES = 10
    x_arr = np.random.rand(NUM_OF_SAMPLES)
    y_arr = np.random.rand(NUM_OF_SAMPLES)
    idx = range(NUM_OF_SAMPLES)


    # In[16]:


    df = pd.DataFrame(list(zip(x_arr, y_arr)), index=idx, columns=["x", "y"])
    df.head()


    # In[20]:


    ax = df.plot.scatter(x="x", y="y", figsize=(12,6))
    fig = ax.get_figure()
    pic_path = PIC_BASE_PATH+"rand_scatter_plot.png"
    fig.savefig(pic_path)

    return pic_path
#END def getRandomScatterPlot():

import seaborn as sns
import sqlite3
def getIrisAnalysisPlot():
    conn = sqlite3.connect(r"c:\Dev\04.Python\06.Office_automation\Powerpoint_presentation\iris_db\iris.sqlite")
    df2 = pd.read_sql_query("SELECT * FROM iris", conn)
    df2.head()


    sns.set()
    sns_plot = sns.pairplot(df2,  
                            hue="Species", height=2.5)
    # sns_plot = sns.pairplot(df2, vars=["SepalWidth", "SepalLength"], 
    #                         hue="Species", height=2.5)


    pic_fname = PIC_BASE_PATH+"iris_analysis.png"
    sns_plot.savefig(pic_fname)
    return pic_fname
#END def getIrisAnalysisPlot():


def getTop20ArtistSales():
    #http://www.sqlitetutorial.net/sqlite-sample-database/

    business_conn = sqlite3.connect(r"c:\Dev\04.Python\06.Office_automation\Powerpoint_presentation\sqlite_sample_db\chinook.db")
    businessdf = pd.read_sql_query("""
        SELECT ArtistName, SUM(Quantity*UnitPrice) as TotalSales
        FROM (
            SELECT * from invoices
            LEFT JOIN invoice_items ON invoices.InvoiceId = invoice_items.InvoiceId
            LEFT JOIN tracks ON invoice_items.TrackId = tracks.TrackID
            LEFT JOIN albums ON albums.AlbumId = tracks.AlbumId
            LEFT JOIN (
                SELECT Name as ArtistName, ArtistId
                FROM artists
            ) as artists2
            ON artists2.ArtistId = albums.ArtistId
        )
        GROUP BY ArtistId
        ORDER BY TotalSales ASC
    """, business_conn)
    # businessdf.head()
    plt1 = businessdf.loc[businessdf.index > len(businessdf)-21].plot.bar(x="ArtistName", 
                                                               y="TotalSales",
                                                                figsize=(10, 8)    )
    # ax = plt.gca()
    plt1.set_ylabel("$")
    plt1.set_xlabel("")
    plt1.set_title("Sales of the top 20 artists")
    plt1.legend().remove()
    # plt.get_figure().subplots_adjust(bottom=0.75)
    plt1.get_figure().tight_layout()
    pic_fname = PIC_BASE_PATH+"top20ArtistSales.png"
    plt1.get_figure().savefig(
        pic_fname,
        bbox_inches = "tight"
    )
    return pic_fname
#END def getTop20ArtistSales():

def getDistributionOfSales():

    business_conn = sqlite3.connect(r"c:\Dev\04.Python\06.Office_automation\Powerpoint_presentation\sqlite_sample_db\chinook.db")
    businessdf = pd.read_sql_query("""
        SELECT ArtistName, SUM(Quantity*UnitPrice) as TotalSales
        FROM (
            SELECT * from invoices
            LEFT JOIN invoice_items ON invoices.InvoiceId = invoice_items.InvoiceId
            LEFT JOIN tracks ON invoice_items.TrackId = tracks.TrackID
            LEFT JOIN albums ON albums.AlbumId = tracks.AlbumId
            LEFT JOIN (
                SELECT Name as ArtistName, ArtistId
                FROM artists
            ) as artists2
            ON artists2.ArtistId = albums.ArtistId
        )
        GROUP BY ArtistId
        ORDER BY TotalSales ASC
    """, business_conn)
    businessdf.head()

    summarized_salesDF = pd.DataFrame(
        businessdf.loc[businessdf.index > len(businessdf)-21]
    )
    # summarized_salesDF.to_clipboard()
    summarized_salesDF.head(30)


    businessdf.loc[businessdf.index <= len(businessdf)-21].sum()[1]


    summarized_salesDF = summarized_salesDF.append({"ArtistName": "Other (with sales below first 20)", 
                    "TotalSales": (businessdf.loc[businessdf.index <= len(businessdf)-21].sum()[1])},
                    ignore_index=True)

    summarized_salesDF.index = summarized_salesDF.ArtistName
    summarized_salesDF.head(35)


    ax = summarized_salesDF.plot.pie("TotalSales", autopct='%1.0f%%', pctdistance=0.7, figsize=(10,10))
    ax.legend().remove()
    ax.set_ylabel("")

    pic_fname = PIC_BASE_PATH+"distribusionOfSales_perArtist.png"
    ax.get_figure().savefig(
        pic_fname,
        bbox_inches = "tight"
    )
    return pic_fname
#END def getDistributionOfSales():

