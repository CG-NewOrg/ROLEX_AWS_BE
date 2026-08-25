namespace devcockpit;

entity PromptTemplates {
  key ID             : UUID;
      Prompt_Details : LargeString;
      Date_Added     : DateTime;
      Category       : String(100);
      MsgType        : String(100);
      UpdatedBy      : String(320);
      UpdatedAt      : DateTime;
      CreatedBy      : String(320);
      Project_Id     : String(320);
      PromptId       : String(1024);
}

entity UserResourceMapping {
  key ID              : Integer @sap.autoIncrement;
      Username        : String(100);
      EmailID         : String(320);
      Project_Details : String(320);
}

entity FileDetails {
  key ID                : Integer  @readonly  @sap.autoIncrement;
      FileName          : String(1024);
      ObjectStoreRefKey : String(1024);
      FileType          : String(1024);
      UpdatedBy         : String(320);
      Category          : String(100);
      UpdatedAt         : DateTime @default: 'CURRENT_TIMESTAMP';
      CreatedBy         : String(320);
      Date_Added        : DateTime @default: 'CURRENT_TIMESTAMP';
      Project           : String(320);
}

entity feedback {
  key Issue_ID    : Integer  @readonly  @sap.autoIncrement;
      IssueTitle  : String(150);
      IssueDetail : String(1024);
      IssueType   : String(100);
      Priority    : String(100);
      IssueStatus : String(45);
      CreatedBy   : String(320);
      CreatedAt   : DateTime @default: 'CURRENT_TIMESTAMP';
}

entity user_login_details {
  key ID               : Integer  @readonly  @sap.autoIncrement;
      Email_Id         : String(320);
      UserName         : String(100);
      login_time       : DateTime;
      logout_time      : DateTime;
      session_duration : Time;
      session_id       : String(100);
      tokens_consumed  : Integer;
      project          : String(320);
}

entity model_usage {
  key ID          : Integer  @readonly  @sap.autoIncrement;
      session_id  : String(100);
      model_id    : String(100);
      model_name  : String(320);
      tokens_used : Integer;
}

entity Prompt_logs {
  key ID             : Integer  @readonly  @sap.autoIncrement;
      user_id        : String(320);
      session_id     : String(100);
      project        : String(320);
      prompt         : LargeString;
      prompt_id      : String(1024);
      system_id      : String(1024);
      sysmsg         : LargeString;
      token_consumed : Integer;
      Date_Added     : DateTime @default: 'CURRENT_TIMESTAMP';
      model_name     : String(320);
}