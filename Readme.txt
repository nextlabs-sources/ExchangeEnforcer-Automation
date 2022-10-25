update
1. support the Exchange2019
2. add the policy deploy and undeploy function, support CC91 or later version

prepare
user prepare:
1. Copy the createuser.ps1 and user.csv to the exchange server
2. Add the users in the user.csv, it's contain username, display name, principal name and department
3. Use the powershell to run the createuser.ps1 script, it's will create users as the user.csv setting. These users all use the 123blue! as default password

CC environment prepare:
1. Login the CC console, and import the Policy for EE autotesting.bin to the CC
2. Modify the CC infomation,open the  EEAuto.exe.config, modify below sections:  CCToolPath and CChost value as your actual environment 
  <appSettings>
    <add key="CCToolPath" value="E:\java\ccAPI\target\ccAPI-1.0-SNAPSHOT-jar-with-dependencies.jar"/>
    <add key="CCHost" value="auto-cc91.auto.com"/>
  </appSettings>
3. Save it.

Exchange Auto testing:
1. Copy the EEAuto to the root C, and rename the EEAuto to the EE
2. Make sure the EE.xml and Test Case.xml in the folder C:/EE
3. Open the EE.xml file, modify below node as your actual environment, the node name, URL,password, FQDN as your actual environment. if you have more than one exchange, add a extra node
	<auto>
		<URL>https://10.23.57.16/EWS/Exchange.asmx</URL>
		<Password>123blue!</Password>
		<FQDN>auto.com</FQDN>
	</auto>
4. Test Case.xml, if you want add more cases, you can open the Test Case.xml, add extra case like the exist case
5. Run the EEAuto.exe
6. Wait the application completed, it's will collect the failed case to the failed_${currenttime}.xml file, you can re-run this failed case by modify the name to Test Case.xml

ps:
1. The CCTool implement the policy deploy and undeploy, and make sure add the "" in the policy name, like the policy name is Automation testing, you should add it fill it as "Automation testing" in the policy seciton
2. The CCTool default password is 12345Next!, make sure when you use the CCTool, change the CC password to 12345Next!