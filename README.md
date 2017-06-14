# yahooFinanceYield
Example how to use yahoo finance package to get yields to excel file (easy to use even if you don't know much about coding like me)

#licence
MIT licence
Just credit me, Riku Pasonen as original source and take note pyhton licenses and MIT licence rules

#


Requirements and insall order:

1) Python 3.X (get the latest)
I reccomend larger package of Anaconda: https://www.continuum.io/downloads

Bare bones version: https://www.python.org/downloads/

Let installer put pyhton to path variable(it asks this in some point)

I tend to prefer 32bit versions but I should not matter in this case(I think)

2) Install some development environment

Anaconda comes with spyder. It ok.

I prefer eclipse:
Tutorial video: https://www.youtube.com/watch?v=hK8Fwelri5Y
(zip comes with eclipse project files also ready)

3) Install packages
Excel package:
from windows command promt type:"pip install openpyxl" (without quotes)
(pip should be installed before but it should come with anaconda at least)

yahoo finance package:

from windows command promt type: "pip install yahoo-finance" (without quotes)


3: Edit the excel stock lists. See yahoo.py file for how to use them in combination. 

Excel file needs to be closed when yahoo.py runs. 
Running yahoo.py is the only thin needed. Should work with double click if pyhton interpreter is assosiated with .py files
I always run with cmd or via the development environment
