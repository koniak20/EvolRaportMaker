# Project Plan

1. ~~Put spiders in default scrapy templete.~~
2. ~~Make each spider to do their job and save data to extern file.~~
3. Make bash/python program which run each spider and then put this to excel.
4. ~~Search for way to transfer data from libfreOffice to 'Word'/ or straightforward to excell~~
5. Transform python file to exe (?). / or embbed this into excel files(?)

## To do

- [x] Test Excel to Libbre, if it is working correctly
- [x] Install scrapy
- [x] Redo/Copy working spiders from old version
- [ ] Spiders
	- [ ] Energy
		- [x] TGE
		- [x] GPI
		- [x] Trans
		- [ ] Pogoda
		- [x] Wiatr/Foto
		- [ ] Prognoza Wiatr
		- [x] Base
		- [ ] Zielone
		- [ ] ss
	- [x] Gas
		- [x] TGE
		- [x] Base
- [ ] Data from spiders to excel
	- [ ] Energy
		- [x] TGE
		- [x] GPI
		- [x] Trans
		- [ ] Pogoda
		- [ ] Prognoza Wiatr
		- [x] Base
		- [ ] Zielone
		- [ ] ss
	- [x] Gas
		- [x] TGE
		- [x] BASE
- [ ] Data from excel to word
	- [ ] Energy
		- [x] TGE
		- [x] GPI
		- [x] Trans
		- [ ] Pogoda
		- [x] Prognoza Wiatr
		- [x] Base
		- [x] Zielone
		- [ ] ss
	- [x] Gas
		- [x] TGE
		- [x] BASE

## Links

- [Scrapy in file](https://stackoverflow.com/questions/21662689/scrapy-run-spider-from-script)
- [Bash Script](https://stackoverflow.com/questions/18686824/running-scrapy-from-a-shell-script)
- [That awesome guy which show you word automation]()

# Notes

- Turning off log in Scrapy *LOG_ENABLED*
- Find the way to transform date to word document
- with 1.2 making raport took 45 min
- Did not find any info about moving excel chart to word, but it is possible to create chart with python and then move it.

### Log
07.06.22

1. Made templete to word GAS and ENERGY.
2. Made first version of gas_trasnfer to word.

05.06.22

1. main.py moving BASE data to excel and also TGEgas data.

03.06.22

1. main.py works as old script

02.06.22

1. Transformed python invoking spiders to windows.

01.06.22

1. Invoking spiders done.

31.05.22 

1. BASEenergy, BASEgas spider done.

30.05.22 

1. Basic bash script which run spiders.

29.05.22 

1. For my use Excel and LibbreOffice seems compatible, nice.
2. Scrapy install.
3. TGEenergy, TGEgas, GPI spider done. 


