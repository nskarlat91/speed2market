## speed2market
Python script for the data manipulation step of S2M Report - see https://confluence.nike.com/display/OBA/Speed+2+Market for detailed documentation

## Installation
Please install numpy, pandas, matplotlib, seaborn

## Input files

- Store Master (Box) --> **C:\Users\NSkarl\Box**\MPO_Replen_PRD\Store Master file.xlsx **PLEASE AMEND TO YOUR LOCAL MACHINE PATH**
- OB --> \\hilversm-nss-01\SharedData03\CustomerSvcs.Countries.EHQ\08.Strategics\04.Nike Retail\7. FACTORY STORES\27-Data Foundation\OB.xlsx"
- SOHER file --> \\hilversm-nss-01\shareddata05\Retail.EHQ\MERCHAND\FACALLOCATION\Allocation analysis\SOH expectation report\Weekly SOH expectation - dist. email.xlsx

## Output files

- speed2market **C:\Users\NSkarl\Box**\Speed 2 Market Dashboard\speed2market.csv **PLEASE AMEND TO YOUR LOCAL MACHINE PATH**

## Data Path
Ensure the data and paths are correct to your local version. The share drive paths are used for most of the files. Please adjust for Store Master input and OB output (those are adjusted to my local machine).

## Important info

Please input store master file from box as it captures all changes that MOLs perform on the capacities of the stores. Please always save the output with the given file name in Box as Tableau uses this Box location as input.

## Usage

python speed2marketv2.py
