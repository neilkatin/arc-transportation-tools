
ARC Transportation tool
=======================

This is a small tool to ease management of the DTT when doing LOG-TRA-SA tasks.

This tool does two different tasks:

## Vehicles with Staff Roster

It annotates entries in the Vehicles excel output from the Disaster Transportation Tool (DTT) with entries
from the Staff Roster report to make following up on missing contracts easier.

The tool only outputs vehicles of type 'R' (rental) and without a Key #.

## Avis Open Rentals  with Vehicles

It takes an Avis Open Rentals report and marks entries that match in the vehicles report.  The color coding is:

* green if all entries match the Avis report to the same entry on the vehicle spreadsheet
* yellow if that entry matches but all don't
* red for the non-matching entries


# Building

## OS pre-build steps

I build and tested with cygwin.  You need these cygwin pages (this example shows installing
them with the
[apt-cyg](https://github.com/transcode-open/apt-cyg) tool, but you can also use the standard cygwin setup program)

(There's no good reason why this shouldn't work on native windows or Linux; I
just didn't test it in those enviroments to get the build environment requirements)

## Python steps

I did all my testing in a virtual environment using python 3.8.  I used the pipenv tool to maintain the environment.

All the pathnames are set in the config.py program.  There are three files to be read:

vehicles.xslx: (required) an excel export from the vehicles tab of the DTT
Staff Roster [date].xlsx: (optional) a current staff roster.  If not present that tab won't be generated
ARC Open Rentals - (optional) mm-dd-yyyy.xlsx: an Avis report.  If not present that tab won't be opened

To run the program:

``` shell
pipenv --python 3.8
pipenv install
pipenv shell
./main.py [ --debug ]
```

by default the program makes a 'merged.xslx' file in the current directory.  This is changable in config.py

