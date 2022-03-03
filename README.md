# Fan Scraping Optimization

## Background

I've been working for an air handling unit manufacturer since 2014. I've been in many departments, but mainly in operations and export in the last few years.

At sales side, we try to make it super simple for the customer to buy our units. Selection software is fast and datasheets provided are good, so "*no problem*" here.

Thing is, the lead time of our machines are somewhat long. We're talking about 8-10 weeks delivery, mostly due to suppliers, so placing purchase orders as early as possible for the right components is critical. In order to do so, technical department must upload the units to the system, and here is where the ball stops rolling.

![Factory](https://raw.githubusercontent.com/darroyolpz/Fan-Scraping-Optimization/master/img/factory.jpg)

It takes around one week-ten days to upload the units, but this component information is key for our suppliers, at least for raw material planning. The fastest way to get the components is getting information from the datasheet, and here is where the scraping algorithm is useful.

## Fan scraping from datasheet

Datasheets are very detailed. This is good for extracting information, but it took me a while to figure out how to do it. This was developed during my days at operations.

I ended up separating each unit using the cover page (so that I know that first unit goes from page 1 to page 10) and then looking for the fan data in that page range.

![First page function](https://raw.githubusercontent.com/darroyolpz/Fan-Scraping-Optimization/master/img/fp_data.jpg)

For extracting fan data information, I relied on keywords in the datasheet, so I could retrieve the information in the middle :smile:

![Fan data](https://raw.githubusercontent.com/darroyolpz/Fan-Scraping-Optimization/master/img/fan_data.jpg)

Now that all that information is stored in a spreadsheet, purchase department can send it to suppliers, negotiate prices and deliveries and so on.

## Faster selections

This second part of the development took place when I started working in export. A customer needed to use our units with a special voltage that was not included in our selection software.

There were 77 units in total, and 24 fan references to choose from. You can expect two minutes per selection, so in total it would have taken around 62 hours. But it gets better: of course customers love to change things in the units, and it's normal to have like 5 or 6 revisions per projects, so we're talking about 310-372 hours to complete this project.

Since I already knew how to extract data from datasheets, I decided to use the Ziehl-Abegg web-service to do this job for me. Thank God I was in ZA factory some years before, so I knew some people to help me with the code.

![ZA factory](https://raw.githubusercontent.com/darroyolpz/Fan-Scraping-Optimization/master/img/za_factory.jpg)

Finally the code was working pretty well and also saved a ton of time for future selections.

## Fans optimization

Our selection program provides solutions for highest energy effiency (we only have like 2 or 3 fans per unit size) but in a cost-driven-market like Spain, Portugal, Greece, etc this might not be the best approach.

While trying to squeeze some units for some customers, I realized we could save up to 8% if we used a different fan than the one provided in our datasheets. Nobody did this exercise before. Doing it manually takes time, but now it was possible with the script.

As a result, we even added more models of fans in a new product we released in 2021. I felt really proud of it, as it all started by trying to facilitate the work of some colleagues :v: