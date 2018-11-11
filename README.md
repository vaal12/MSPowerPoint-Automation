# Microsoft PowerPoint presentation from SQL database


## What this does
The script "main.py"
1. Takes "template" PowerPoint presentation - "Template presentation.pptx"
2. Populates it with data which is being taken on the fly from the chinook.db. 
3. Saves presentation as "Template presentation_[date].pptx" so the template presentation stays intact
4. Creates a PDF file from the saved presentation "Template presentation_2_[date].pdf"

Resulting files are: 
- [Template presentation_11-11-2018.pptx](https://raw.githubusercontent.com/vaal12/MSPowerPoint-Automation/master/Template%20presentation_11-11-2018.pptx)
- [Template presentation_11-11-2018.pdf](https://raw.githubusercontent.com/vaal12/MSPowerPoint-Automation/master/Template%20presentation_11-11-2018.pdf)

Those 2 files are also in this repository.

## Data Source

Sqlite database is obtained from http://www.sqlitetutorial.net/sqlite-sample-database/. 

Database schema: ![sqlite sample DB Schema](https://raw.githubusercontent.com/vaal12/MSPowerPoint-Automation/master/sqlite_sample_db/sqlite-sample-database-color.jpg)

DB file for convenience is also located at sqlite_sample_db\chinook.db in this repository

## Prerequisities to run the script

Script requires Python 3.x, pandas, numpy and matplotlib installed to pythons installation.
Anaconda from https://www.anaconda.com/download/ was used as dev environment and so script should be run from it as:


    C:\Anaconda3\python.exe  C:\Anaconda3\cwp.py C:\Anaconda3 C:\Anaconda3\python.exe c:\Dev\04.Python\06.Office_automation\Powerpoint_presentation\main.py 



## Notes 

As the database sctructure is quite normalized, queries to get aggregate data require some skill.
As example see below a query which is needed for getting top 20 artists (by sales)

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

## License
Unless required by applicable law or agreed to in writing, software distributed under the License is distributed on an "AS IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
