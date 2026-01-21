**TL;DR** - A Python automation script for tablets to self-heal a freezing clock-in and out web app. It auto-starts on boot, launches the card-reader app and a clean Selenium browser, monitors crashes and internet connectivity via PIDs, and restarts only affected apps instead of rebooting tablets. It also injects custom JS to fix input-focus bugs and periodically refresh the page, eliminating frequent physical IT intervention.

<details>
A python project aimed for self-fixing automation of tablets behaviour for a ship company that was released in the production fleet wide. The problem was, at the time, is that whenever an employee would clock in and out with their personal ID (employee card), after some time, the tablet web app would freeze/not respond, web app page would go completely white or the input box would have the wrong focus and an IT employee would need to manually go and reboot the tablet which was a hassle since they're mounted on the walls with metal frames and taking them out is not as easy. The idea behind this automation is to elimiate the frequent manual intervention of IT employees and let the script take care of it.
>
>The script starts running automatically on boot, runs the necessarry app (PCSC Keyboard) for the card reader to work and opens Selenium webdriver browser instance (clean browser with no cache or any user settings) that loads the web page for clocking in and out. The app needs to be opened first before the browser otherwise the reader won't insert the numbers in the input field. The terminal in the background would display general status as info of the tablet. PID is used to track if the PCSC Keyboard or web crashes, if they do, the app that did not crash would be closed with the given PID and then all the apps would be restarted instead of an entire tablet being rebooted. The same applies for internet connectivity status, the script will continuously check for internet connection and once it has been reestablished, apps will be restarted in order. Additionally for some reason the devs of that web page didn't add input field re-focus so if somebody manually types their employee ID and PIN to clock in or out, and the page returns, the input focus will remain on PIN so when the next person that comes and uses the card reader, their employee ID will be inserted in the PIN input instead of employee ID input, to fix that small yet often nuisance, the script also injects custom JS that takes care for refocusing on page return as well as page refresh every 1 min.
</details>

<img width="1481" height="1009" alt="autotar" src="https://github.com/user-attachments/assets/12ea0bc5-fd24-49b1-ba0b-1e15e7e3fbf7" />

---
Title: AutoTAR  
Desc: Automation and enchancments of TAR tablet behaviour  
Ver: 1.1  

Last ver fixes and enchancments:
- added pre-disabled quick edit for console pause-state prevention
- added task scheduler schedule info from sys32com
- added better text status of the current settings
- added URL handler in case URL is wrong

