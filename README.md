* This project is built on AppScript. It is used to aggregate risk profile data entry files and provide details on incomplete department records, sending corresponding reminder emails. Admin file contains 3 main functions:
  * It is set up with triggers to run the function triggerImportMHSRR(), which fetches data from raw files, and the function triggerImportMSRR(), which aggregates incomplete risk profile codes, at 08:00 daily.
  * With an email sending frequency of 3 times per month, the triggerGuiEmail() function will run at 09:00 daily, checking the email sending date condition and sending the corresponding form.
  * btnCapNhatQuyen() function sharing with multiple people, multiple files.

* Workflow: 
![Picture1](https://github.com/lovegreen21/Risk-Assessment-Report/assets/129417444/45700007-04b7-4ea9-91aa-dcb2d31b7652)

* Sheet contains file link:

  ![Picture2](https://github.com/lovegreen21/Risk-Assessment-Report/assets/129417444/19fc95b8-6358-4ddf-b42b-74c4ff37d51d)

* Sheet contains raw data after aggregating:
![Picture3](https://github.com/lovegreen21/Risk-Assessment-Report/assets/129417444/0bc673cd-e5ff-4b31-98d9-c5619b692fd3)

* Email Layout (Note: Personal information is hidden):
 ![Picture4](https://github.com/lovegreen21/Risk-Assessment-Report/assets/129417444/7bec324a-591a-4c6d-8385-95ce5968410f)




  
