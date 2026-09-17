using devcockpit as devCockpit from '../db/schema';


service cockpit @(requires: 'Viewer') {
    entity user_login_details  as projection on devCockpit.user_login_details;
    entity model_usage         as projection on devCockpit.model_usage;
    entity UserResourceMapping as projection on devCockpit.UserResourceMapping;

    action   getProjectDetailsOfUser(payload: projectDetails)                                              returns String;
    // action   getDeployments(User_Email_Id: String)                                                         returns String;
    function getPromptDetails(Category: String, MsgType: String, ProjectId: String, scenario: String)      returns String;
    function getFiles(Category: String, Project: String)                                                   returns String;
    action   saveLogin(payload: loginDetails)                                                              returns String;
    action   saveLogout(payload: logoutDetails)                                                            returns String;
    function getLoginDetails(Email_Id: String, project: String)                                            returns String;
    function getLoginDetailsOfAllUserAd(project: String)                                                     returns String;
    function extract_docxAzure(url: String)                                                                returns String;
    action   createFeedback(payload: feedbackDetails)                                                      returns String;
    action   createPromptDetails(payload: PromptInput)                                                     returns String;
    action   deletePromptDetails(uuid: String)                                                             returns String;
    action   logTokenUsage(payload: tokendetails)                                                          returns String;
    action   getPromptDetailsofUser2_0(project: String, user_id: String, date_added: String)               returns String;
    action   updateProject(session_id: String, project: String)                                            returns String;
    action   uploadFile(payload: upload)                                                                   returns String;
    function getFileDetails(key: String)                                                                   returns LargeString;
    action   deleteFiles(files: array of String)                                                           returns String;
    action   deleteFilesFromKB(kb: Boolean, filenames: array of String, category: String, project: String) returns String;
    function sessionDataExcelAd(fromDate: String, toDate: String)                                            returns LargeString;
    function promptDataExcel(project: String, fromDate: String, toDate: String)                            returns LargeString;
    action   generateDocument(content: LargeString, templateKey: String, tabName: String);
    function viewTemplate(key: String)                                                                     returns String;
    function extractTemplateStructure(key: String)                                                          returns String;
    function getAllUsers()                                                                                 returns array of UserResourceMapping;
    function getAllFeedback()                                                                              returns String;
    action   addUser(payload: userAddition)                                                                returns UserResourceMapping;
    action   deleteUser(emailId: String(320))                                                              returns Boolean;
    action   updateUser(payload: userUpdate)                                                               returns UserResourceMapping;
    // action   uploadUsersExcel(fileBase64: LargeString)                                                     returns String;
    action   readFileFromGit(payload: path)                                                                returns String;
    action   getGitRepoTreeStructure(payload: branch)                                                      returns String;
    action   getAllBranches(payload: allBranch)                                                            returns String;
    action   pushFileToGit(payload: gitPayload)                                                            returns String;
    type upload {
        Category   : String;
        Project    : String;
        userId     : String;
        fileName   : String;
        mimeType   : String;
        fileBase64 : LargeString;
    }

    type projectDetails {
        UserId   : String;
        userName : String;
    }

    type tokendetails {
        session_id      : String;
        model_id        : String;
        model_name      : String;
        tokensGenerated : Integer;
        user_id         : String;
        Prompt          : String;
        sysmsg          : LargeString;
        date_added      : String;
        system_id       : String;
        project         : String;
    }

    type userAddition {
        username       : String(100);
        emailId        : String(320);
        projectDetails : String(320);
    }

    type userUpdate {
        id             : Integer;
        emailId        : String(320);
        newEmailId     : String(320);
        username       : String(100);
        projectDetails : String(320);
    }

    type PromptInput {
        Prompt_Details : String;
        Category       : String;
        MsgType        : String;
        PromptId       : String;
        ProjectId      : String;
        UserId         : String;
        DateTime       : String;
    }

    type feedbackDetails {
        IssueTitle  : String;
        IssueDetail : String;
        Priority    : String;
        IssueType   : String;
        UserId      : String;
        DateTime    : String;
    }

    type loginDetails {
        Email_Id   : String;
        UserName   : String;
        login_time : DateTime;
    }

    type logoutDetails {
        Email_Id    : String;
        UserName    : String;
        logout_time : DateTime;
        session_id  : String;
    }
      type path {
        pathAccess : String;
        branchName : String;
        repo       : String;
        token      : String;
        username   : String;
    }

    type branch {
        branchName : String;
        repo       : String;
        token      : String;
        username   : String;
    }

    type gitPayload {
        filePath     : String;
        content      : String;
        commitMsg    : String;
        userName     : String;
        emailId      : String;
        branchName   : String;
        targetBranch : String;
        repo         : String;
        token        : String;
        username     : String;
    }

    type allBranch {
        repo     : String;
        username : String;
        token    : String
    }


}
