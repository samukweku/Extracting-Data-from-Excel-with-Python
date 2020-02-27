### This script shows how to extract multiple sheets from multiple files.<br>


```python
#import libraries

from pathlib import Path
from fnmatch import fnmatch
import pandas as pd
```


```python
#This extracts only the excel files

folder = Path('004-MSPTDA-ExcelFiles')

all_files = list(folder.iterdir())

#fnmatch uses pattern matching to get the files we are interested in
#the name method in pathlib returns a string, of the file name
#the string method lower converts text to lowercase 
#this allows us to search for all xls* files present

excel_only_files = [x for x in folder.iterdir()
                    if fnmatch(x.name.lower(),'*.xls*')
                   ]
```


```python
def process_files(xls_file):
    '''
    Picks an excel file,
    extracts the city name using the stem method from Path,
    filters out any sheet name that starts with Sheet,
    reads in the excel file using pandas, with the specified sheet names,
    concatenates the resulting dictionary,
    assigns city back to the dataframe
    and returns a dataframe
    '''
    
    xls = pd.ExcelFile(xls_file)
    
    #stem is a pathlib method
    #that gets just the name of the file
    #with the parent and suffix stripped off
    city = xls_file.stem 
    
    sheet_list = [sheet for sheet in xls.sheet_names                  
                  if not sheet.startswith('Sheet')]
    
    df = pd.read_excel(xls,sheet_list)
    
    outcome = pd.concat([data.assign(SalesPerson = SalesPerson) 
                         for SalesPerson, data in df.items()],
                        ignore_index=True
                       ).assign(City = city)
    
    return outcome    
```


```python
#applies function to each excel file
combo = [process_files(xls_file) for xls_file in excel_only_files]

#lumps all the processed files into one dataframe
outcome = pd.concat(combo, ignore_index=True)
```


```python
outcome.sample(20)
```




<div>
<style scoped>
    .dataframe tbody tr th:only-of-type {
        vertical-align: middle;
    }

    .dataframe tbody tr th {
        vertical-align: top;
    }

    .dataframe thead th {
        text-align: right;
    }
</style>
<table border="1" class="dataframe">
  <thead>
    <tr style="text-align: right;">
      <th></th>
      <th>Date</th>
      <th>Product</th>
      <th>Units</th>
      <th>Sales</th>
      <th>SalesPerson</th>
      <th>City</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <th>33159</th>
      <td>2016-12-08</td>
      <td>Majestic Beaut</td>
      <td>144</td>
      <td>4852.80</td>
      <td>Chin</td>
      <td>Portland</td>
    </tr>
    <tr>
      <th>9978</th>
      <td>2017-01-04</td>
      <td>Bellen</td>
      <td>132</td>
      <td>3986.40</td>
      <td>Sindy</td>
      <td>Seattle</td>
    </tr>
    <tr>
      <th>5018</th>
      <td>2016-07-17</td>
      <td>Majestic Beaut</td>
      <td>96</td>
      <td>3321.60</td>
      <td>Pham</td>
      <td>Tacoma</td>
    </tr>
    <tr>
      <th>10374</th>
      <td>2017-12-21</td>
      <td>Sunshine</td>
      <td>132</td>
      <td>2824.80</td>
      <td>Sindy</td>
      <td>Seattle</td>
    </tr>
    <tr>
      <th>47395</th>
      <td>2016-11-21</td>
      <td>Carlota</td>
      <td>84</td>
      <td>2520.00</td>
      <td>Popi</td>
      <td>Oakland</td>
    </tr>
    <tr>
      <th>32617</th>
      <td>2016-12-14</td>
      <td>Aussie Round</td>
      <td>216</td>
      <td>7026.48</td>
      <td>Chin</td>
      <td>Portland</td>
    </tr>
    <tr>
      <th>17153</th>
      <td>2016-12-18</td>
      <td>Majestic Beaut</td>
      <td>48</td>
      <td>1721.28</td>
      <td>Gigi</td>
      <td>Seattle</td>
    </tr>
    <tr>
      <th>18465</th>
      <td>2017-12-06</td>
      <td>Bellen</td>
      <td>156</td>
      <td>4924.92</td>
      <td>Gigi</td>
      <td>Seattle</td>
    </tr>
    <tr>
      <th>48351</th>
      <td>2016-12-31</td>
      <td>Bellen</td>
      <td>108</td>
      <td>3204.36</td>
      <td>Popi</td>
      <td>Oakland</td>
    </tr>
    <tr>
      <th>2658</th>
      <td>2017-11-07</td>
      <td>Carlota</td>
      <td>108</td>
      <td>3149.28</td>
      <td>Pham</td>
      <td>Tacoma</td>
    </tr>
    <tr>
      <th>42689</th>
      <td>2017-01-27</td>
      <td>Aussie Round</td>
      <td>84</td>
      <td>2707.32</td>
      <td>Fran</td>
      <td>Oakland</td>
    </tr>
    <tr>
      <th>43493</th>
      <td>2017-12-04</td>
      <td>Sunshine</td>
      <td>60</td>
      <td>1313.40</td>
      <td>Fran</td>
      <td>Oakland</td>
    </tr>
    <tr>
      <th>40920</th>
      <td>2017-12-22</td>
      <td>Quad</td>
      <td>96</td>
      <td>4013.76</td>
      <td>Fran</td>
      <td>Oakland</td>
    </tr>
    <tr>
      <th>43406</th>
      <td>2016-11-25</td>
      <td>Quad</td>
      <td>180</td>
      <td>8168.40</td>
      <td>Fran</td>
      <td>Oakland</td>
    </tr>
    <tr>
      <th>16782</th>
      <td>2017-10-28</td>
      <td>Tri Fly</td>
      <td>72</td>
      <td>396.00</td>
      <td>Gigi</td>
      <td>Seattle</td>
    </tr>
    <tr>
      <th>9337</th>
      <td>2017-11-28</td>
      <td>Majestic Beaut</td>
      <td>108</td>
      <td>3944.16</td>
      <td>Sindy</td>
      <td>Seattle</td>
    </tr>
    <tr>
      <th>35481</th>
      <td>2017-12-29</td>
      <td>Tri Fly</td>
      <td>84</td>
      <td>483.84</td>
      <td>Chin</td>
      <td>Portland</td>
    </tr>
    <tr>
      <th>48502</th>
      <td>2017-11-06</td>
      <td>Aussie Round</td>
      <td>120</td>
      <td>3876.00</td>
      <td>Popi</td>
      <td>Oakland</td>
    </tr>
    <tr>
      <th>3500</th>
      <td>2017-10-22</td>
      <td>Bellen</td>
      <td>240</td>
      <td>7080.00</td>
      <td>Pham</td>
      <td>Tacoma</td>
    </tr>
    <tr>
      <th>35144</th>
      <td>2017-11-24</td>
      <td>Bellen</td>
      <td>144</td>
      <td>4235.04</td>
      <td>Chin</td>
      <td>Portland</td>
    </tr>
  </tbody>
</table>
</div>


