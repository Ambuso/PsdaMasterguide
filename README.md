# Master Guide Examinations 2026

Online Church Heritage and Bible Truths examinations for the Master Guide class.
Instructor: Ambuso Dismas.

Candidates open the site, choose a paper, enter their name and local church, and have 45 minutes to answer 60 questions.
The paper marks itself, and each candidate sends their result to the instructor on WhatsApp.

Live site: https://ambuso.github.io/PsdaMasterguide/

## Results and emails

`apps-script/Exam_Results_Receiver.gs` is the Google Apps Script that records every submitted paper in the
"Master Guide Exam Results 2026" Google Sheet and emails each candidate one message with both results after
they finish the second paper. To update it: paste it over the code at script.google.com, save, then
Deploy → Manage deployments → pencil → Version: New version → Deploy.
