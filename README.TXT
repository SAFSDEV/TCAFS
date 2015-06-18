
This TCAFS Project has been "upgraded" to TestComplete V9.

It is part of and distributed with the SAFS Test Automation Framework. 

If you are running an earlier version of TestComplete you will need
to upgrade to TestComplete 9 or transfer the Script and UserForms 
to your TestComplete Project repository.  Alternatively, you might be 
able to overwrite this V9 repository with an appropriate earlier 
version.


To use\verify the SAFS/TCAFS Engine:
====================================

1. Test Complete 9.0 or higher must be installed and
   connecting properly to any local or remote License Server.
   
2. If not done as part of the original SAFS install,  OR
   If Test Complete has been installed or upgraded AFTER the SAFS 
   install:
   
   a. Execute C:\SAFS\SetupTCAFS.wsf, which will:
      
      I. Set the %TESTCOMPLETE_HOME% environment variable to the
         correct Test Complete install directory.

3. Verify SAFS/TCAFS runtime integration:

   a. Execute C:\SAFS\Project\runTCAFSTest.bat
   
      I. SAFS, Test Complete, and a sample test should be seen.
      
   b. View the SAFS log output at:
   
      I. C:\SAFS\Project\Datapool\Logs\TCAFSCycle.txt
      
   c. SAFS and Test Complete should shutdown upon test completion.
   

4. Problem launching Test Complete in Step #3 above?

   a. Verify\Edit the contents of C:\TCAFS\TCAFS.vbs:
   
      I.  'projectname' and 'suitename' point to existing SAFS assets.
      II. 'command' points to correct and existing executable:
   
          i.\bin\TestComplete.exe, or \bin\TestExecute.exe

   b. Try Step #3 again after all edits are complete.
   