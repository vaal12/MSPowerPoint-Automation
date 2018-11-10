# Project Title

Microsoft PowerPoint presentation creation from SQL (sqlite) database. Buisness analytics with pandas and graphing with matpotlib

## Data Source

Sqlite database is obtained from #http://www.sqlitetutorial.net/sqlite-sample-database/. 

Database schema: ![sqlite sample DB Schema](http://www.sqlitetutorial.net/wp-content/uploads/2015/11/sqlite-sample-database-color.jpg)

DB file for convenience is also located at sqlite_sample_db\chinook.db in this repository

### Result

The script creates 
1. Takes "template" PowerPoint presentation - "Template presentation.pptx" and 2
2. populates it with data which is being taken on the fly from the chinook.db. 
3. Saves presentation as "Template presentation_2_[date].pptx" so the template presentation stays intact
4. Creates a PDF file from the saved presentation "Template presentation_2_[date].pdf"

Resulting files are: 
- Template presentation_2_11-11-2018.pptx
- Template presentation_2_11-11-2018.pdf
Also stored in this repository


```
Give examples
```

### Installing

A step by step series of examples that tell you how to get a development env running

Say what the step will be

```
Give the example
```

And repeat

```
until finished
```

End with an example of getting some data out of the system or using it for a little demo

## Running the tests

Explain how to run the automated tests for this system

### Break down into end to end tests

Explain what these tests test and why

```
Give an example
```

### And coding style tests

Explain what these tests test and why

```
Give an example
```

## Deployment

Add additional notes about how to deploy this on a live system

## Built With

* [Dropwizard](http://www.dropwizard.io/1.0.2/docs/) - The web framework used
* [Maven](https://maven.apache.org/) - Dependency Management
* [ROME](https://rometools.github.io/rome/) - Used to generate RSS Feeds

## Contributing

Please read [CONTRIBUTING.md](https://gist.github.com/PurpleBooth/b24679402957c63ec426) for details on our code of conduct, and the process for submitting pull requests to us.

## Versioning

We use [SemVer](http://semver.org/) for versioning. For the versions available, see the [tags on this repository](https://github.com/your/project/tags). 

## Authors

* **Billie Thompson** - *Initial work* - [PurpleBooth](https://github.com/PurpleBooth)

See also the list of [contributors](https://github.com/your/project/contributors) who participated in this project.

## License

This project is licensed under the MIT License - see the [LICENSE.md](LICENSE.md) file for details

## Acknowledgments

* Hat tip to anyone whose code was used
* Inspiration
* etc
