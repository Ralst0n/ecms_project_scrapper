# ecms_project_scrapper
Gets latest proposed projects from ECMS

Each week using bs4 and Selenium webdriver This program logs into ECMS using my credentials and checks to see what projects are being proposed so our project managers
can see what options are coming up and (if necessary) find partners to prepare to go after projects they like. Rather than having someone remember
to log in and check the projects, then verify the project is a type of work that the company does and isn't something that has been passed around already.

![image of the selenium webdriver and the filter algorithm in action](http://res.cloudinary.com/ralst0n/image/upload/v1510071837/Scrapper_running_ghbngd.jpg)

ecms project scrapper handles all of that for us. Each project that is of the correct type (in PA that's just Construction inspection) is added to a list and written to an excel file.
after all proposed projects are filtered by python that new list is checked against the old list and any projects that were already in the old list are removed. 
finally an excel macro formats the new list into an HTML email that is sent out to all of the management in Pennsylvania.

![image of the email sent to PA managers each week](http://res.cloudinary.com/ralst0n/image/upload/v1510071837/ECMS_email_cclrdj.jpg)


The program runs automatically at lunch time each week thanks to windows task scheduler
