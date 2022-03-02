# Fan Scraping Optimization

## Background

Ok so here's the deal. I've been working for an air handling unit manufacturer since 2014. I've been in many departments, but mainly in operations and export in the last few years.

At sales side, we try to make it super simple for the customer to buy our units. Selection software is fast and datasheets provided are good, so "*no problem*" here.

Thing is, the lead time of our machines are somewhat long. We're talking about 8-10 weeks delivery, mostly due to suppliers, so placing purchase orders as early as possible for the right components is critical. In order to do so, technical department must upload the units to the system, and here is where the ball stops rolling.

![Factory](https://raw.githubusercontent.com/darroyolpz/Fan-Scraping-Optimization/master/img/factory.jpg)

It takes around one week-ten days to upload the units, but this component information is key for our suppliers, at least for raw material planning. The fastest way to get the components is getting information from the datasheet, and here is where the scraping algorithm is useful.

## Fan Scraping

Datasheets are very detailed. This is good for extracting information, but it took me a while to figure out how to do it.

I ended up separating each unit using the cover page (so that I know that first unit goes from page 1 to page 10) and then looking for the fan data in that page range.

![First page function](https://raw.githubusercontent.com/darroyolpz/Fan-Scraping-Optimization/master/img/fp_data.jpg)

```python
# First page function -------------------------------------
def fpFunction():
   print('Starting first page function--------------------')
   inner_list, outter_list = [], []
   for page, pageEnd in zip(aPageStart, aPageEnd):
      print('Looking at', page, 'page')
      pageContent = extractContent(page)
      print('\n')

      # Get line
      wordStart = 'Unit name:'
      wordEnd = 'Fecha'
      line = get_value_function(pageContent, wordStart, wordEnd)

      # Reset ahu_value for each pageStart
      ahu_value = []

      # Get AHU type
      for ahu in ahus:
         if ahu in pageContent:
            ahu_value.append(ahu)

      # In case of DV10 or DV100, always get the longest one
      if len(ahu_value) == 1:
         ahu = ahu_value[0]
      elif len(ahu_value) > 1:
         print('Possible conflict!')
         final_ahu = ''
         for value in ahu_value:
            if len(value) > len(final_ahu):
               final_ahu = value
         ahu = final_ahu
         print('Final value:', ahu)
         print('\n')
      else:
         ahu = '---'

      # Get reference
      wordStart = 'Planta no. '
      wordEnd = 'Unit'
      ref = get_value_function(pageContent, wordStart, wordEnd)

      # Airflow
      wordStart = ')'
      wordEnd = 'm'
      airflow = get_value_function(pageContent, wordStart, wordEnd)

      inner_list = [page, pageEnd, line, ahu, ref, airflow]
      outter_list.append(inner_list)

   return outter_list
```

For extracting fan data information, I relied on keywords in the datasheet, so I could retrieve the information in the middle :)

![Fan data](https://raw.githubusercontent.com/darroyolpz/Fan-Scraping-Optimization/master/img/fan_data.jpg)

```python
def extractFeatures(aWordStart, aWordEnd, pageStart, pageEnd, allowed_pages = 1):
   outter_list = []
   for page in range(pageStart, pageEnd):
      # Initiate the inner_list and get the page number
      inner_list = []
      
      # Extract page content
      pageContent = extractContent(page)
      print('Checking at page number', page+1)

      for wordStart, wordEnd in zip(aWordStart, aWordEnd):
         print('Looking for ', wordStart, 'and', wordEnd)

         # Work in starting and ending pairs, page by page
         if (wordStart in pageContent) and (wordEnd in pageContent):
            print('Found on page', page+1)
            print('\n')
            print(pageContent)
            print('\n')
            unitFeature = get_value_function(pageContent, wordStart, wordEnd)

            # Important in case the next wordStart is above the previos one
            print('Feature found:', unitFeature)
            split_word = unitFeature + wordEnd
            print('Split_word:', split_word)
            if split_word in pageContent:
               print('Split_word in pageContent')
               print('\n')
            posEnd = pageContent.index(split_word)
            pageContent = pageContent[posEnd:]

            if unitFeature == 'Error flag!':
               print('Error flag! Length not correct.')
               break
            else:
               inner_list.append(unitFeature)

         # If cheking of additional pages is allowed
         elif allowed_pages > 0:
            if len(inner_list) == 0:
               # Reset inner list
               inner_list = []
               print('No luck this time')
               print('\n')
               # Exit loop and go for the next page
               break
            elif len(inner_list) > 0:
               until_page = page + allowed_pages + 1
               for new_page in range(page + 1, until_page):
                  pageContent = extractContent(new_page)
                  try:
                     print('\n')
                     print('New page being checked')
                     unitFeature = get_value_function(pageContent, wordStart, wordEnd)
                     inner_list.append(unitFeature)
                  except:
                     print('No luck even in the next page')
                     print('\n')
                     # Exit loop and go for the next page
                     break

      # Check the lenght and append to the outter list
      if len(inner_list) == len(aWordStart):
         print('New entry for the outter list!')
         print('\n')
         # Add the number of page in the inner list
         inner_list = [page + 1, *inner_list] # In order to show real page number
         outter_list.append(inner_list)
         # Reset inner list for next feature
         inner_list = []

   try:
      return outter_list
   except:
      print('No outter_list found!')
```