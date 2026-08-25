
const cds = require('@sap/cds');
const mammoth = require('mammoth');
const axios = require("axios");
const path = require('path');
const { PutObjectCommand, ListObjectsV2Command, GetObjectCommand, HeadObjectCommand, DeleteObjectsCommand, CopyObjectCommand } = require('@aws-sdk/client-s3');
const ExcelJS = require('exceljs');
const { Paragraph, TextRun, Table, TableCell, TableRow, WidthType, HeadingLevel, Document, Packer, TableOfContents } = require('docx');
const PizZip = require('pizzip'); const { DOMParser, XMLSerializer } = require('@xmldom/xmldom');
const { getObjectStoreConfig } = require('./config/aws');
require('dotenv').config();
const { v4: uuidv4 } = require('uuid');
const { getDestination } = require('@sap-cloud-sdk/connectivity');
const { executeHttpRequest } = require('@sap-cloud-sdk/http-client');


async function aggregateModelUsageBySession(data) {
  const sessionMap = {};

  data.forEach(row => {
    const sessionId = row.session_id;

    if (!sessionMap[sessionId]) {
      sessionMap[sessionId] = {
        session_id: sessionId,
        login_time: row.login_time,
        logout_time: row.logout_time,
        session_duration: row.session_duration,
        tokens_consumed: row.tokens_consumed,
        project: row.project,
        Email_Id: row.Email_Id,
        UserName: row.UserName,
        models: []
      };
    }

    sessionMap[sessionId].models.push({
      model_name: row.model_name,
      tokens_used: row.tokens_used
    });
  });

  return Object.values(sessionMap);
}

const containsEmoji = (text) => {
  const emojiRegex = /[\u{1F600}-\u{1F64F}]|[\u{1F300}-\u{1F5FF}]|[\u{1F680}-\u{1F6FF}]|[\u{1F1E0}-\u{1F1FF}]|[\u{2600}-\u{26FF}]|[\u{2700}-\u{27BF}]|[\u{1F900}-\u{1F9FF}]|[\u{1FA00}-\u{1FA6F}]|[\u{1FA70}-\u{1FAFF}]/gu;
  return emojiRegex.test(text);
};

const containsSpecialChars = (text) => {
  const allowedPattern = /^[\w\s.,;:!?'"()\-[\]{}@#$%&*+=<>/\\|`~\n\r\t]*$/;
  return !allowedPattern.test(text);
};

const isValidEmail = (email) => {
  const regex =
    /^[A-Za-z0-9!#$%&'*+/=?^_`{|}~-]+(\.[A-Za-z0-9!#$%&'*+/=?^_`{|}~-]+)*@[A-Za-z0-9-]+(\.[A-Za-z0-9-]+)*\.[A-Za-z]{2,}$/;

  if (!regex.test(email)) return false;

  const [local, domain] = email.split("@");

  if (local.includes("..") || domain.includes("..")) return false;

  if (
    local.startsWith(".") ||
    local.endsWith(".") ||
    domain.startsWith(".") ||
    domain.endsWith(".")
  ) return false;

  const domainParts = domain.split(".");
  for (const part of domainParts) {
    if (part.startsWith("-") || part.endsWith("-")) {
      return false;
    }
  }
  return true;
};

const isNullOrNA = (val) =>
  val === null || val === undefined || val === '' ||
  (typeof val === 'string' && (val.trim().toLowerCase() === 'null' || val.trim().toUpperCase() === 'NA'));

function mergeModelUsage(modelsArray, newModels) {
  if (!newModels || !Array.isArray(newModels)) return;
  newModels.forEach(({ model_name, tokens_used }) => {
    const existing = modelsArray.find(m => m.model_name === model_name);
    if (existing) {
      existing.tokens_used += (tokens_used || 0);
    } else {
      modelsArray.push({ model_name, tokens_used: tokens_used || 0 });
    }
  });
}

function cleanCellValue(cell) {
  if (cell.value) {
    const val = String(cell.value).trim();
    if (
      ['null', 'undefined', 'null(null)', 'null(0)'].includes(val.toLowerCase()) ||
      /^null\(\d*\)$/.test(val) ||
      /\(0\)$/.test(val)
    ) {
      cell.value = '';
    }
  }
}

async function uploadToObjectStore(objectStoreRefKey, buffer, mimeType) {
  const { containerUri, sasToken } = await getObjectStoreConfig();
  const blobUrl = `${containerUri}/${encodeURIComponent(objectStoreRefKey)}`;
  const signedUrl = `${blobUrl}?${sasToken}`;

  await axios.put(signedUrl, buffer, {
    headers: {
      "x-ms-blob-type": "BlockBlob",
      "Content-Type": mimeType
    }
  });
}


module.exports = cds.service.impl(async function () {

  this.on('createPromptDetails', async (req) => {
    try {
      const { Prompt_Details, Category, MsgType, ProjectId, PromptId, UserId, DateTime } = req.data.payload;

      if (isNullOrNA(MsgType)) return req.reject(400, 'Please provide Message Type');
      if (isNullOrNA(Category)) return req.reject(400, 'Please provide Category');
      if (isNullOrNA(ProjectId)) return req.reject(400, 'Please provide ProjectId');
      if (ProjectId === 'default') return req.reject(400, 'Project cannot be default');
      if (isNullOrNA(Prompt_Details)) return req.reject(400, 'Please provide Message');
      if (isNullOrNA(UserId)) return req.reject(400, 'Please provide UserId');
      if (!isValidEmail(UserId)) return req.reject(400, 'UserId must be a valid email address');

      // if (req.user?.id?.toLowerCase() !== UserId?.toLowerCase()) {
      //   return req.reject(400, 'Mismatched email id');
      // }

      if (MsgType === 'sysMsg') {
        if (isNullOrNA(PromptId)) return req.reject(400, 'Please provide System Message Name');
      }

      let generatedPromptId = PromptId;
      if (Prompt_Details && typeof Prompt_Details === 'string') {
        if (containsEmoji(Prompt_Details)) {
          return req.reject(400, `${MsgType === 'sysMsg' ? 'System Message' : 'Prompt'} contains emojis. Please provide only text and numbers.`);
        }
        if (containsSpecialChars(Prompt_Details)) {
          return req.reject(400, `${MsgType === 'sysMsg' ? 'System Message' : 'Prompt'} contains invalid special characters.`);
        }
      }

      const insertNewRecord = async () => {
        const newId = uuidv4();
        const insertResult = await cds.db.tx(async tx => {
          return tx.run(
            INSERT.into('devcockpit_PromptTemplates').entries({
              ID: newId,
              Prompt_Details: Prompt_Details,
              Date_Added: DateTime || new Date().toISOString(),
              Category: Category,
              MsgType: MsgType,
              CreatedBy: UserId,
              Project_Id: ProjectId,
              PromptId: generatedPromptId
            })
          );
        });
        return { status: 201, message: 'Record inserted successfully', data: insertResult, PromptId: generatedPromptId };
      };

      if (MsgType === 'prompt') {
        const existingPrompt = await cds.db.tx(async tx => {
          return tx.run(
            SELECT.from('devcockpit_PromptTemplates')
              .columns('id')
              .where(`prompt_details = '${Prompt_Details}' AND category = '${Category}' AND msgtype = '${MsgType}' AND (project_id = '${ProjectId}' OR project_id = 'default')`)
          );
        });

        if (existingPrompt.length === 0) {
          if (PromptId) {
            const existingById = await cds.db.tx(async tx => {
              return tx.run(
                SELECT.from('devcockpit_PromptTemplates')
                  .columns('id', 'project_id', 'category')
                  .where(`promptid = '${PromptId}' AND (project_id = '${ProjectId}' OR project_id = 'default')`)
              );
            });

            if (existingById.length > 0) {
              // Use lowercase column names when accessing returned data
              const defaultRecord = existingById.find(r => r.project_id === 'default');
              if (defaultRecord) {
                return req.reject(403, `Cannot modify PromptId '${PromptId}' - it is a default ${MsgType} which is read-only.`);
              } else if (existingById[0].project_id === ProjectId && existingById[0].category === Category) {
                const updateResult = await cds.db.tx(async tx => {
                  return tx.run(
                    UPDATE('devcockpit_PromptTemplates')
                      .set({
                        Prompt_Details: Prompt_Details,
                        UpdatedAt: DateTime,
                        UpdatedBy: UserId
                      })
                      .where({ id: existingById[0].id })
                  );
                });
                return { status: 200, message: 'Record updated successfully', data: updateResult };
              } else {
                return req.reject(403, 'Given PromptId does not exist in this project/category, cannot update');
              }
            } else {
              generatedPromptId = PromptId;
              return await insertNewRecord();
            }
          } else {
            const countResult = await cds.db.tx(async tx => {
              return tx.run(
                SELECT.from('devcockpit_PromptTemplates').columns({ nextid: { func: 'count', args: ['*'] } })
              );
            });
            const maxId = (Number(countResult[0]?.nextid) || 0) + 1;
            generatedPromptId = `${Category}_${ProjectId}_${maxId}_prompt`;
            return await insertNewRecord();
          }
        } else {
          return req.reject(403, 'Detailed prompt already exists');
        }
      }

      if (MsgType === 'sysMsg') {
        const existingSysMsg = await cds.db.tx(async tx => {
          return tx.run(
            SELECT.from('devcockpit_PromptTemplates')
              .columns('id', 'project_id')
              .where(`promptid = '${generatedPromptId}' AND (project_id = '${ProjectId}' OR project_id = 'default')`)
          );
        });

        if (existingSysMsg.length > 0) {
          // Use lowercase column names when accessing returned data
          if (existingSysMsg[0].project_id === 'default') {
            return req.reject(403, 'Cannot update default system message');
          }
          const updateResult = await cds.db.tx(async tx => {
            return tx.run(
              UPDATE('devcockpit_PromptTemplates')
                .set({
                  Prompt_Details: Prompt_Details,
                  UpdatedAt: DateTime,
                  UpdatedBy: UserId
                })
                .where({
                  category: Category,
                  msgtype: MsgType,
                  promptid: generatedPromptId,
                  project_id: ProjectId
                })
            );
          });
          return { status: 200, message: 'Record updated successfully', data: updateResult };
        } else {
          return await insertNewRecord();
        }
      }
    } catch (err) {
      console.error('Error in createPromptDetails:', err);
      return req.reject(500, `Error processing request: ${err.message}`);
    }
  });

  this.on('deletePromptDetails', async (req) => {
    try {
      if (!req.user.is('Admin')) {
        return req.error(403, 'Access denied. Admin role required.');
      }
      const { uuid } = req.data;

      if (!uuid) {
        return req.reject(400, 'Please provide uuid to delete');
      }

      const existingRecords = await cds.db.tx(async tx => {
        return tx.run(
          SELECT.from('devcockpit_PromptTemplates')
            .columns('ID', 'Project_Id')
            .where({ ID: uuid })
        );
      });

      if (!existingRecords || existingRecords.length === 0) {
        return req.reject(404, `Record with uuid '${uuid}' not found`);
      }

      if (existingRecords[0].Project_Id === 'default') {
        return req.reject(403, 'Cannot delete record – default records are read-only');
      }

      await cds.db.tx(async tx => {
        return tx.run(
          DELETE.from('devcockpit_PromptTemplates').where({ ID: uuid })
        );
      });

      return {
        status: 200,
        message: 'Record deleted successfully',
        deletedUuid: uuid
      };

    } catch (err) {
      console.error('Error in deletePromptDetails:', err);
      return req.reject(500, `Error processing request: ${err.message}`);
    }
  });


  this.on('getPromptDetails', async (req) => {
    try {
      const { Category, MsgType, ProjectId, scenario } = req.data;

      if (!ProjectId) {
        return req.reject(400, 'ProjectId is required');
      }

      const categoryFilter = Category || scenario;
      const msgTypeFilter = MsgType || 'prompt';

      console.log('getPromptDetails called with:', { Category, scenario, MsgType, ProjectId, categoryFilter, msgTypeFilter });

      let queryResult;
      if (categoryFilter) {
        queryResult = await cds.db.tx(async tx => {
          return tx.run(
            SELECT.from('devcockpit_PromptTemplates')
              .columns(
                'ID', 'Prompt_Details', 'Date_Added', 'Category', 'MsgType',
                'UpdatedBy', 'UpdatedAt', 'CreatedBy', 'Project_Id', 'PromptId'
              )
              .where(`MsgType = '${msgTypeFilter}' AND Category = '${categoryFilter}' AND (Project_Id = '${ProjectId}' OR Project_Id = 'default')`)
              .orderBy({ Date_Added: 'desc' })
          );
        });
      } else {
        queryResult = await cds.db.tx(async tx => {
          return tx.run(
            SELECT.from('devcockpit_PromptTemplates')
              .columns(
                'ID', 'Prompt_Details', 'Date_Added', 'Category', 'MsgType',
                'UpdatedBy', 'UpdatedAt', 'CreatedBy', 'Project_Id', 'PromptId'
              )
              .where(`MsgType = '${msgTypeFilter}' AND (Project_Id = '${ProjectId}' OR Project_Id = 'default')`)
              .orderBy({ Date_Added: 'desc' })
          );
        });
      }

      console.log('Query returned rows:', queryResult?.length || 0);

      const processedResults = (queryResult || []).map(row => {
        if (row.Prompt_Details && Buffer.isBuffer(row.Prompt_Details)) {
          return { ...row, Prompt_Details: row.Prompt_Details.toString('utf-8') };
        }
        return row;
      });

      return {
        status: 200,
        result: processedResults,
        count: processedResults.length,
        message: processedResults.length === 0 ? 'No Data found' : 'Executed'
      };

    } catch (err) {
      console.error('Error in getPromptDetails:', err);
      return req.error({
        code: '500',
        message: `Internal Server error: ${err.message}`,
        target: 'getPromptDetails',
        status: 500
      });
    }
  });



  const fetchAIModels = async (sqlResponse, req) => {
    try {

      const jwt = req.headers?.authorization?.split(' ')[1];
      if (!jwt) {
        console.warn('fetchAIModels: No JWT found in Authorization header');
      }
      let destination;
      try {
        destination = await getDestination({
          destinationName: 'AI_CORE_CGAI_COCKPIT',
          jwt: jwt
        });
      } catch (destError) {
        console.error('Error fetching destination AI_Core:', {
          message: destError.message,
          cause: destError.cause?.message,
          rootCause: destError.cause?.cause?.message,
          stack: destError.stack
        });
        throw new Error(`Destination retrieval failed: ${destError.message}`);
      }

      if (!destination) {
        throw new Error('AI_Core destination not found');
      }

      console.log('Destination AI_Core retrieved successfully:', {
        name: destination.name,
        url: destination.url,
        authenticationType: destination.authentication,
        hasOriginalProperties: !!destination.originalProperties,
        originalPropertiesKeys: destination.originalProperties ? Object.keys(destination.originalProperties) : []
      });
      console.log("destination", destination)
      const apiVersion = destination.originalProperties?.APIVERSION || destination.originalProperties?.destinationConfiguration.APIVERSION || null;

      let response;
      try {
        console.log('Calling AI Core API: GET /lm/deployments');
        response = await executeHttpRequest(
          {
            destinationName: 'AI_CORE_CGAI_COCKPIT',
            jwt: jwt
          },
          {
            method: 'GET',
            url: '/lm/deployments',
            headers: {
              'AI-Resource-Group': 'default',
              'Content-Type': 'application/json'
            }
          }
        );
        console.log('AI Core API response status:', response.status);
      } catch (httpError) {
        console.error('AI Core API request failed:', {
          message: httpError.message,
          cause: httpError.cause?.message,
          rootCause: httpError.cause?.cause?.message,
          stack: httpError.stack,
          responseData: httpError.response?.data,
          responseStatus: httpError.response?.status
        });

        if (httpError.message?.includes('Failed to build headers')) {
          throw new Error(
            'Failed to build authentication headers. Please verify the AI_Core destination has valid OAuth2 credentials (clientId, clientSecret, tokenServiceURL) in SAP BTP Destination Service.'
          );
        }
        if (httpError.message?.toLowerCase().includes('token')) {
          throw new Error(
            `OAuth token retrieval failed: ${httpError.message}. Please check the destination OAuth2 configuration and ensure the token service URL is accessible.`
          );
        }
        if (httpError.response?.status === 401 || httpError.response?.status === 403) {
          throw new Error(
            `Authentication/Authorization failed (HTTP ${httpError.response.status}): ${httpError.response?.data?.message || httpError.message}. Check AI Core credentials and permissions.`
          );
        }
        throw new Error(`AI Core API request failed: ${httpError.message}`);
      }

      const resources = response?.data?.resources || [];
      console.log("AI Core deployments retrieved, count:", resources.length);

      const rearrangedArr = [];
      const gpt4oArr = [];

      resources.forEach((ele) => {
        const hasBackendDetails =
          ele.details &&
          ele.details.resources &&
          ele.details.resources.backendDetails &&
          Object.keys(ele.details.resources.backendDetails).length > 0;

        if (!hasBackendDetails || ele.targetStatus !== 'RUNNING') {
          return;
        }

        if (ele.configurationName === 'GPT_4o') {
          gpt4oArr.push(ele);
        } else if (
          ele.configurationName.includes('GPT') &&
          !ele.configurationName.includes('text')
        ) {
          rearrangedArr.unshift(ele);
        } else if (!ele.configurationName.includes('text')) {
          rearrangedArr.push(ele);
        }
      });

      if (gpt4oArr.length > 0) {
        rearrangedArr.unshift(gpt4oArr[0]);
      }

      return {
        deployments: rearrangedArr,
        sqlResponse: {
          EMAILID: sqlResponse.EmailID,
          APIVERSION: apiVersion,
          PROJECT_DETAILS: sqlResponse.Project_Details
        }
      };

    } catch (e) {
      console.error('fetchAIModels error:', {
        message: e.message,
        stack: e.stack
      });
      req.error({
        code: '500',
        message: `Internal Server Error: ${e.message}`,
        target: 'fetchAIModels',
        status: 500
      });
    }
  };

  this.on('getFileDetails', async (req) => {
    try {
      const { key } = req.data;

      if (!key) {
        return req.error(400, 'key parameter is required');
      }

      const { s3, bucketName } = await getObjectStoreConfig();
      const options = {
        Bucket: bucketName,
        Key: key
      };


      const getCommand = new GetObjectCommand(options);
      const getObjectData = await s3.send(getCommand);

      const ext = path.extname(key).toLowerCase();
      const chunks = [];
      for await (const chunk of getObjectData.Body) {
        chunks.push(chunk);
      }
      const buffer = Buffer.concat(chunks);

      if (ext === '.docx') {
        const result = await mammoth.extractRawText({ buffer });
        return result.value;
      } else if (ext === '.txt') {
        return buffer.toString('utf8');
      } else {
        return buffer.toString('base64');
      }
    } catch (error) {
      if (error.name === 'NotFound') {
        return req.error(404, 'Object not found');
      }
      console.error('Error in getFileDetails:', error);
      return req.error(500, 'Internal Server Error: ' + error.message);
    }
  });


  
  this.on('extract_docxAzure', async (req) => {
    const { buffer } = req.file;
    try {
      const text = await mammoth.extractRawText({ buffer });
      return text.value;
    } catch (error) {
      console.error('Error during extraction from URL:', error);
      req.error(500, 'Extraction failed from the provided URL');
    }
  });


  this.on('deleteFiles', async (req) => {

    const { files } = req.data;

    try {
      if (!req.user.is('Admin')) {
        return req.error(403, 'Access denied. Admin role required.');
      }
      const { s3, bucketName } = await getObjectStoreConfig();

      if (!files || files.length === 0) {
        return req.error(400, 'No files provided for deletion');
      }

      const objectsToDelete = files.map((file) => ({ Key: file }));
      const deleteParams = {
        Bucket: bucketName,
        Delete: {
          Objects: objectsToDelete,
        },
      };


      const entity = 'devcockpit.FileDetails';
      const queryResult = await cds.run(
        DELETE.from(entity)
          .where({
            ObjectStoreRefKey: { in: files }
          })
      );


      const rowsDeleted = queryResult.rowCount || queryResult;
      if (rowsDeleted > 0) {
        const deleteCommand = new DeleteObjectsCommand(deleteParams);
        const deleteResponse = await s3.send(deleteCommand);
        return JSON.stringify({ message: "Selected files deleted successfully", deleteResponse });
      } else {
        return JSON.stringify({ message: "Files does not exist in system, please recheck" });
      }
    } catch (error) {
      console.error("Error deleting files:", error);
      return req.error(500, "Error deleting files: " + error.message);
    }
  });

  this.on('deleteFilesFromKB', async (req) => {
    try {
      if (!req.user.is('Admin')) {
        return req.error(403, 'Access denied. Admin role required.');
      }

      const { kb, filenames, category, project } = req.data;
      console.log('[deleteFilesFromKB] Incoming payload:', { kb, filenames, category, project });

      if (!filenames || filenames.length === 0) {
        return req.error(400, 'No filenames provided for deletion');
      }
      if (!category) {
        return req.error(400, 'category is required');
      }
      if (!project) {
        return req.error(400, 'project is required');
      }

      const jwt = req.headers?.authorization?.split(' ')[1];
      console.log('[deleteFilesFromKB] JWT present:', !!jwt);

      let destination;
      try {
        destination = await getDestination({
          destinationName: 'RagQueryAPI',
          jwt: jwt
        });
        console.log('[deleteFilesFromKB] Destination resolved:', {
          name: destination?.name,
          url: destination?.url,
          authType: destination?.authentication
        });
      } catch (destErr) {
        console.error('[deleteFilesFromKB] Failed to resolve destination:', destErr.message, destErr.stack);
        return req.error(500, 'Failed to resolve KB_integration destination: ' + destErr.message);
      }
      if (!destination) {
        return req.error(500, 'KB_integration destination not found');
      }

      const payload = {
        KB: kb,
        filenames: filenames,
        category: category,
        project: project
      };
      console.log('[deleteFilesFromKB] Sending payload to KB:', JSON.stringify(payload));

      let response;
      try {
        response = await executeHttpRequest(
          {
            destinationName: 'RagQueryAPI',
            jwt: jwt
          },
          {
            method: 'POST',
            url: '/DeleteFromObjectStore',
            headers: {
              'Content-Type': 'application/json'
            },
            data: payload
          }
        );
        console.log('[deleteFilesFromKB] KB response status:', response.status);
        console.log('[deleteFilesFromKB] KB response data:', JSON.stringify(response.data));
      } catch (httpErr) {
        console.error('[deleteFilesFromKB] HTTP request to KB failed:', {
          message: httpErr.message,
          responseStatus: httpErr.response?.status,
          responseData: JSON.stringify(httpErr.response?.data),
          stack: httpErr.stack
        });
        return req.error(500, 'KB HTTP request failed: ' + httpErr.message);
      }

      return JSON.stringify({
        message: 'Files deleted from KB successfully',
        status: response.status,
        data: response.data
      });

    } catch (error) {
      console.error('[deleteFilesFromKB] Unexpected error:', {
        message: error.message,
        stack: error.stack,
        responseStatus: error.response?.status,
        responseData: JSON.stringify(error.response?.data)
      });
      return req.error(500, 'Error deleting files from KB: ' + error.message);
    }
  });


  this.on("generateDocument", async (req) => {

    try {

      let { content, templateKey, tabName } = req.data;

      if (!content) {

        return req.error(400, "content is required");

      }



      content = content.replace(/```[a-zA-Z]*\n?/g, "").replace(/```/g, "").replace(/\r/g, "").trim();



      function sanitizeMarkdown(text) {



        return text

          .replace(/\*\*(.*?)\*\*/g, "$1")

          .replace(/\*(.*?)\*/g, "$1");



      }



      async function streamToBuffer(stream) {

        return new Promise((resolve, reject) => {

          const chunks = [];

          stream.on("data", chunk => chunks.push(chunk));

          stream.once("error", reject);

          stream.once("end", () => resolve(Buffer.concat(chunks)));

        });

      }



      // Extract default styles (font, color, size) from the template's word/styles.xml

      function extractDefaultStyles(zipObj) {

        const defaults = {

          defaultFont: "Calibri",

          defaultSize: 24,

          defaultColor: "000000",

          heading1Font: "Calibri",

          heading1Size: 32,

          heading1Color: null,

          heading2Font: "Calibri",

          heading2Size: 28,

          heading2Color: null,

          heading3Font: "Calibri",

          heading3Size: 26,

          heading3Color: null,

          heading4Font: "Calibri",

          heading4Size: 24,

          heading4Color: null,

          labelColor: null,

          tableFill: "D9E1F2"

        };

        try {

          // ─── Parse theme colors from word/theme/theme1.xml ───

          const themeColors = {};

          try {

            const themeFile = zipObj.file("word/theme/theme1.xml");

            if (themeFile) {

              const themeXml = themeFile.asText();

              const themeDoc = new DOMParser().parseFromString(themeXml);

              // Parse a:themeElements > a:clrScheme

              const clrSchemes = themeDoc.getElementsByTagName("a:clrScheme");

              if (clrSchemes.length > 0) {

                const clrScheme = clrSchemes[0];

                // Standard theme color names and their XML element names

                const themeColorMap = {

                  "dk1": "a:dk1",

                  "lt1": "a:lt1",

                  "dk2": "a:dk2",

                  "lt2": "a:lt2",

                  "accent1": "a:accent1",

                  "accent2": "a:accent2",

                  "accent3": "a:accent3",

                  "accent4": "a:accent4",

                  "accent5": "a:accent5",

                  "accent6": "a:accent6",

                  "hlink": "a:hlink",

                  "folHlink": "a:folHlink"

                };

                for (const [themeName, elementName] of Object.entries(themeColorMap)) {

                  const elements = clrScheme.getElementsByTagName(elementName);

                  if (elements.length > 0) {

                    const colorElement = elements[0];

                    // Check for srgbClr (direct hex color)

                    const srgbClr = colorElement.getElementsByTagName("a:srgbClr");

                    if (srgbClr.length > 0) {

                      const val = srgbClr[0].getAttribute("val");

                      if (val) themeColors[themeName] = val;

                    } else {

                      // Check for sysClr (system color with lastClr fallback)

                      const sysClr = colorElement.getElementsByTagName("a:sysClr");

                      if (sysClr.length > 0) {

                        const lastClr = sysClr[0].getAttribute("lastClr");

                        if (lastClr) themeColors[themeName] = lastClr;

                      }

                    }

                  }

                }

              }

              console.log("[generateDocument] Theme colors resolved:", JSON.stringify(themeColors));

            }

          } catch (themeErr) {

            console.warn("[generateDocument] Could not parse theme1.xml:", themeErr.message);

          }



          // Helper: resolve a color value considering theme references

          function resolveColor(colorElement) {

            if (!colorElement) return null;

            const val = colorElement.getAttribute("w:val");

            const themeColor = colorElement.getAttribute("w:themeColor");

            // If themeColor attribute exists, try to resolve from theme

            if (themeColor && themeColors[themeColor]) {

              let resolved = themeColors[themeColor];

              // Apply themeShade/themeTint if present

              const themeShade = colorElement.getAttribute("w:themeShade");

              const themeTint = colorElement.getAttribute("w:themeTint");

              if (themeShade) {

                resolved = applyShade(resolved, themeShade);

              } else if (themeTint) {

                resolved = applyTint(resolved, themeTint);

              }

              return resolved;

            }

            // Direct hex color

            if (val && val !== "auto") return val;

            return null;

          }



          // Apply shade (darken): multiply each RGB component by shade/255

          function applyShade(hexColor, shadeHex) {

            const shade = parseInt(shadeHex, 16) / 255;

            const r = Math.round(parseInt(hexColor.substring(0, 2), 16) * shade);

            const g = Math.round(parseInt(hexColor.substring(2, 4), 16) * shade);

            const b = Math.round(parseInt(hexColor.substring(4, 6), 16) * shade);

            return (

              r.toString(16).padStart(2, "0") +

              g.toString(16).padStart(2, "0") +

              b.toString(16).padStart(2, "0")

            ).toUpperCase();

          }



          // Apply tint (lighten): tint each component towards 255

          function applyTint(hexColor, tintHex) {

            const tint = parseInt(tintHex, 16) / 255;

            const r = Math.round(parseInt(hexColor.substring(0, 2), 16) * (1 - tint) + 255 * tint);

            const g = Math.round(parseInt(hexColor.substring(2, 4), 16) * (1 - tint) + 255 * tint);

            const b = Math.round(parseInt(hexColor.substring(4, 6), 16) * (1 - tint) + 255 * tint);

            return (

              Math.min(255, r).toString(16).padStart(2, "0") +

              Math.min(255, g).toString(16).padStart(2, "0") +

              Math.min(255, b).toString(16).padStart(2, "0")

            ).toUpperCase();

          }



          const stylesFile = zipObj.file("word/styles.xml");

          if (!stylesFile) return defaults;

          const stylesXml = stylesFile.asText();

          const stylesDoc = new DOMParser().parseFromString(stylesXml);



          // Extract default run properties from w:docDefaults

          const docDefaults = stylesDoc.getElementsByTagName("w:docDefaults");

          if (docDefaults.length > 0) {

            const rPrDefault = docDefaults[0].getElementsByTagName("w:rPrDefault");

            if (rPrDefault.length > 0) {

              const rPr = rPrDefault[0].getElementsByTagName("w:rPr");

              if (rPr.length > 0) {

                const rFonts = rPr[0].getElementsByTagName("w:rFonts");

                if (rFonts.length > 0) {

                  const ascii = rFonts[0].getAttribute("w:ascii") || rFonts[0].getAttribute("w:asciiTheme");

                  if (ascii && !ascii.includes("Theme")) {

                    defaults.defaultFont = ascii;

                  } else if (rFonts[0].getAttribute("w:asciiTheme")) {

                    // Resolve font from theme if possible

                    const themeRef = rFonts[0].getAttribute("w:asciiTheme");

                    // majorHAnsi/minorHAnsi map to theme fonts - check theme

                    try {

                      const themeFile = zipObj.file("word/theme/theme1.xml");

                      if (themeFile) {

                        const themeXml = themeFile.asText();

                        const themeDoc = new DOMParser().parseFromString(themeXml);

                        if (themeRef && themeRef.includes("minor")) {

                          const minorFonts = themeDoc.getElementsByTagName("a:minorFont");

                          if (minorFonts.length > 0) {

                            const latin = minorFonts[0].getElementsByTagName("a:latin");

                            if (latin.length > 0) {

                              const typeface = latin[0].getAttribute("typeface");

                              if (typeface) defaults.defaultFont = typeface;

                            }

                          }

                        } else if (themeRef && themeRef.includes("major")) {

                          const majorFonts = themeDoc.getElementsByTagName("a:majorFont");

                          if (majorFonts.length > 0) {

                            const latin = majorFonts[0].getElementsByTagName("a:latin");

                            if (latin.length > 0) {

                              const typeface = latin[0].getAttribute("typeface");

                              if (typeface) defaults.defaultFont = typeface;

                            }

                          }

                        }

                      }

                    } catch (e) { /* ignore font theme resolution failure */ }

                  }

                }

                const sz = rPr[0].getElementsByTagName("w:sz");

                if (sz.length > 0) {

                  const val = parseInt(sz[0].getAttribute("w:val"), 10);

                  if (val) defaults.defaultSize = val;

                }

                const color = rPr[0].getElementsByTagName("w:color");

                if (color.length > 0) {

                  const resolved = resolveColor(color[0]);

                  if (resolved) defaults.defaultColor = resolved;

                }

              }

            }

          }



          // Extract heading styles

          const styles = stylesDoc.getElementsByTagName("w:style");

          for (let i = 0; i < styles.length; i++) {

            const style = styles[i];

            const styleId = style.getAttribute("w:styleId") || "";

            const type = style.getAttribute("w:type") || "";

            if (type !== "paragraph") continue;



            let headingLevel = 0;

            if (/^Heading1$|^heading1$/i.test(styleId) || styleId === "Heading1") headingLevel = 1;

            else if (/^Heading2$|^heading2$/i.test(styleId) || styleId === "Heading2") headingLevel = 2;

            else if (/^Heading3$|^heading3$/i.test(styleId) || styleId === "Heading3") headingLevel = 3;

            else if (/^Heading4$|^heading4$/i.test(styleId) || styleId === "Heading4") headingLevel = 4;



            if (headingLevel === 0) {

              // Check outlineLvl

              const outlineLvl = style.getElementsByTagName("w:outlineLvl");

              if (outlineLvl.length > 0) {

                const lvl = parseInt(outlineLvl[0].getAttribute("w:val"), 10);

                if (lvl >= 0 && lvl <= 3) headingLevel = lvl + 1;

              }

            }



            if (headingLevel >= 1 && headingLevel <= 4) {

              const rPr = style.getElementsByTagName("w:rPr");

              if (rPr.length > 0) {

                const rFonts = rPr[0].getElementsByTagName("w:rFonts");

                if (rFonts.length > 0) {

                  const ascii = rFonts[0].getAttribute("w:ascii");

                  if (ascii) {

                    defaults[`heading${headingLevel}Font`] = ascii;

                  } else {

                    // Try to resolve from theme font reference

                    const asciiTheme = rFonts[0].getAttribute("w:asciiTheme");

                    if (asciiTheme) {

                      try {

                        const themeFile = zipObj.file("word/theme/theme1.xml");

                        if (themeFile) {

                          const themeXml = themeFile.asText();

                          const themeDoc = new DOMParser().parseFromString(themeXml);

                          if (asciiTheme.includes("minor")) {

                            const minorFonts = themeDoc.getElementsByTagName("a:minorFont");

                            if (minorFonts.length > 0) {

                              const latin = minorFonts[0].getElementsByTagName("a:latin");

                              if (latin.length > 0) {

                                const typeface = latin[0].getAttribute("typeface");

                                if (typeface) defaults[`heading${headingLevel}Font`] = typeface;

                              }

                            }

                          } else if (asciiTheme.includes("major")) {

                            const majorFonts = themeDoc.getElementsByTagName("a:majorFont");

                            if (majorFonts.length > 0) {

                              const latin = majorFonts[0].getElementsByTagName("a:latin");

                              if (latin.length > 0) {

                                const typeface = latin[0].getAttribute("typeface");

                                if (typeface) defaults[`heading${headingLevel}Font`] = typeface;

                              }

                            }

                          }

                        }

                      } catch (e) { /* ignore */ }

                    }

                  }

                }

                const sz = rPr[0].getElementsByTagName("w:sz");

                if (sz.length > 0) {

                  const val = parseInt(sz[0].getAttribute("w:val"), 10);

                  if (val) defaults[`heading${headingLevel}Size`] = val;

                }

                const color = rPr[0].getElementsByTagName("w:color");

                if (color.length > 0) {

                  const resolved = resolveColor(color[0]);

                  if (resolved) {

                    defaults[`heading${headingLevel}Color`] = resolved;

                    if (headingLevel === 4) defaults.labelColor = resolved;

                  }

                }

              }

            }

          }



          // If heading colors were not explicitly defined, use the document's default color

          for (let h = 1; h <= 4; h++) {

            if (!defaults[`heading${h}Color`]) {

              defaults[`heading${h}Color`] = defaults.defaultColor;

            }

          }



          // Set labelColor from heading4 or heading3, or fall back to defaultColor

          defaults.labelColor = defaults.heading4Color || defaults.heading3Color || defaults.defaultColor;



          // Propagate default font to headings that didn't specify their own

          for (let h = 1; h <= 4; h++) {

            if (defaults[`heading${h}Font`] === "Calibri" && defaults.defaultFont !== "Calibri") {

              defaults[`heading${h}Font`] = defaults.defaultFont;

            }

          }



          return defaults;

        } catch (e) {

          console.warn("[generateDocument] Could not parse template styles, using defaults:", e.message);

          return defaults;

        }

      }







      function splitNumberedLabel(text) {

        const m = text.match(/^(\d+\s*[).\\-]?\s*[^:]+:)(.*)$/);

        if (!m) return null;

        return { prefix: m[1], remainder: m[2] || "" };

      }



      function splitBoldBeforeColon(text, withDash = false, fontName) {

        const idx = text.indexOf(":");

        if (idx === -1) {

          return [new TextRun({ text: (withDash ? "- " : "") + text, font: fontName })];

        }

        return [

          new TextRun({

            text: (withDash ? "- " : "") + text.slice(0, idx + 1),

            bold: true,

            font: fontName

          }),

          new TextRun({ text: text.slice(idx + 1), font: fontName })

        ];

      }



      function normalizeBullet(text) {

        return text.replace(/^[•●∙‣▸►▪]\s*/, "- ");

      }



      function borders() {

        return {

          top: { style: "single", size: 1, color: "000000" },

          bottom: { style: "single", size: 1, color: "000000" },

          left: { style: "single", size: 1, color: "000000" },

          right: { style: "single", size: 1, color: "000000" }

        };

      }



      function mergeBrokenTableRows(rows) {

        const merged = [];

        let currentRow = "";

        rows.forEach((line, index) => {

          const cleaned = line.replace(/\t/g, "|").trim();

          if (/^[\s\-|:]+$/.test(cleaned)) return;

          if (index === 0) {

            merged.push(cleaned);

            return;

          }

          const pipeCount = (cleaned.match(/\|/g) || []).length;

          if (pipeCount >= 2) {

            if (currentRow) merged.push(currentRow);

            currentRow = cleaned;

          } else {

            currentRow += " | " + cleaned;

          }

        });

        if (currentRow) merged.push(currentRow);

        return merged;

      }



      function parseMarkdownTable(rows) {

        if (!rows || rows.length < 2) return null;

        rows = mergeBrokenTableRows(rows);

        const splitRow = (row) =>

          row.replace(/^\|/, "").replace(/\|$/, "").split("|").map(c => c.trim());

        const header = splitRow(rows[0]);

        const body = rows

          .slice(1)

          .filter(r => !/^[\-\s|:]+$/.test(r))

          .filter(r => r.replace(/\|/g, "").trim().length > 0)

          .map(splitRow);

        const colCount = header.length;

        const normalizedBody = body.map(row => {

          if (row.length < colCount) {

            return [...row, ...Array(colCount - row.length).fill("")];

          }

          return row.slice(0, colCount);

        });

        return new Table({

          width: { size: 100, type: WidthType.PERCENTAGE },

          rows: [

            new TableRow({

              tableHeader: true,

              children: header.map(cell =>

                new TableCell({

                  shading: {

                    fill: "D9E1F2" // light blue like FSD

                  },

                  borders: borders(),

                  children: [

                    new Paragraph({

                      children: [

                        new TextRun({

                          text: cell,

                          bold: true

                        })

                      ]

                    })

                  ]

                })

              )

            }),

            ...normalizedBody.map(row =>

              new TableRow({

                children: row.map(cell =>

                  new TableCell({

                    borders: borders(),

                    children: [

                      new Paragraph({

                        children: [new TextRun({ text: sanitizeMarkdown(cell) })]

                      })

                    ]

                  })

                )

              })

            )

          ]

        });

      }



      function templateHasTOC(templateDoc) {

        const instrTexts = templateDoc.getElementsByTagName("w:instrText");

        for (let i = 0; i < instrTexts.length; i++) {

          const text = instrTexts[i].textContent || "";

          if (text.includes("TOC")) {

            return true;

          }

        }

        return false;

      }

      function safeAppend(parent, node, ref) {

        const imported = templateDoc.importNode(node, true);



        if (ref && ref.parentNode === parent) {

          parent.insertBefore(imported, ref);

        } else if (ref) {



          const sectPrElement = parent.getElementsByTagName("w:sectPr")[0];

          if (sectPrElement && sectPrElement.parentNode === parent) {

            parent.insertBefore(imported, sectPrElement);

          } else {

            parent.appendChild(imported);

          }

        } else {

          parent.appendChild(imported); // fallback

        }

      }



      // ─── Fetch template FIRST so we can extract its styles ───

      const { s3, bucketName } = await getObjectStoreConfig();

      const s3Response = await s3.send(

        new GetObjectCommand({

          Bucket: bucketName,

          Key: templateKey

        })

      );

      const templateBuffer = await streamToBuffer(s3Response.Body);

      const templateZip = new PizZip(templateBuffer);



      // Extract template styles (font, color, size) for consistent formatting

      const tplStyles = extractDefaultStyles(templateZip);

      console.log("[generateDocument] Template styles extracted:", JSON.stringify(tplStyles));



      const parser = new DOMParser();

      const templateDoc = parser.parseFromString(templateZip.file("word/document.xml").asText());

      // ─── End template fetch ───



      const children = [];

      let insideTable = false;

      let currentTableRows = [];



      const lines = content.split("\n");

      lines.forEach((rawText, index) => {

        if (!rawText) return;

        let text = sanitizeMarkdown(rawText.trim());

        if (!text) return;

        const nextLine = lines[index + 1]?.trim() || "";

        if (/^##\s*SECTION\s+\d+/i.test(text)) {

          children.push(new Paragraph({

            heading: HeadingLevel.HEADING_1,

            spacing: { before: 400, after: 200 },

            children: [

              new TextRun({

                text: text.replace(/^##\s*/, ""),

                bold: true,

                color: tplStyles.heading1Color,

                size: tplStyles.heading1Size,

                font: tplStyles.heading1Font

              })

            ]

          }));

          return;

        }

        if (insideTable && !(text.includes("|") && text.split("|").length >= 3)) {



          const table = parseMarkdownTable(currentTableRows);



          if (table) {

            children.push(

              new Paragraph({ text: "" }),

              table,

              new Paragraph({ text: "" })

            );

          } else {

            currentTableRows.forEach(r => {

              children.push(new Paragraph({ text: r }));

            });

          }



          currentTableRows = [];

          insideTable = false;

        }

        // if (/^(?:\*\*)?\s*SECTION\s+\d+.*$/i.test(text)) {

        if (/^(?:##\s*)?(?:\*\*)?\s*SECTION\s+\d+.*$/i.test(text)) {

          const clean = text

            .replace(/^##\s*/, "")

            .replace(/\*\*/g, "")

            .replace(/━+/g, "")

            .replace(/\s+/g, " ")

            .trim();



          children.push(new Paragraph({

            heading: HeadingLevel.HEADING_1,

            spacing: { before: 300, after: 150 },

            children: [

              new TextRun({

                text: clean,

                bold: true,

                color: tplStyles.heading1Color,

                size: tplStyles.heading1Size,

                font: tplStyles.heading1Font

              })

            ]

          }));



          return;

        }

        if (/^━+\s*SECTION\s+\d+/i.test(text)) {



          const clean = text

            .replace(/━+/g, "")

            .replace(/\s*—\s*/g, " — ")   // normalize dash

            .replace(/\s+/g, " ")

            .trim();



          children.push(new Paragraph({

            heading: HeadingLevel.HEADING_1,

            spacing: { before: 400, after: 200 },

            children: [

              new TextRun({

                text: clean,

                bold: true,

                color: tplStyles.heading1Color,

                size: tplStyles.heading1Size,

                font: tplStyles.heading1Font

              })

            ]

          }));



          return;

        }

        if (text.startsWith("#### ")) {

          const headingText = text.slice(5);



          children.push(new Paragraph({

            outlineLevel: 3, // Level 4 heading (0-indexed)

            spacing: { before: 150, after: 100 },

            children: [new TextRun({ text: headingText, bold: true, size: tplStyles.heading4Size, font: tplStyles.heading4Font, color: tplStyles.heading4Color })]

          }));

          return;

        }

        if (text.startsWith("### ")) {

          const isMainSection = /^###\s*\d+.\s+/.test(text);

          const headingText = text.replace(/^###\s*/, "");

          // Use outlineLevel instead of heading style to avoid template's automatic numbering

          children.push(new Paragraph({

            outlineLevel: isMainSection ? 0 : 2, // Level 1 or 3 heading (0-indexed)

            spacing: { before: 300, after: 150 },

            children: [new TextRun({

              text: headingText,

              bold: true,

              color: isMainSection ? tplStyles.heading1Color : tplStyles.heading3Color,

              size: isMainSection ? tplStyles.heading1Size : tplStyles.heading3Size,

              font: isMainSection ? tplStyles.heading1Font : tplStyles.heading3Font

            })]

          }));

          return;

        }

        if (text.startsWith("## ")) {

          const headingText = text.slice(3);

          // Use outlineLevel instead of heading style to avoid template's automatic numbering

          children.push(new Paragraph({

            outlineLevel: 1, // Level 2 heading (0-indexed)

            spacing: { before: 250, after: 120 },

            children: [new TextRun({ text: headingText, bold: true, color: tplStyles.heading2Color, size: tplStyles.heading2Size, font: tplStyles.heading2Font })]

          }));

          return;

        }

        if (text.startsWith("# ")) {

          const headingText = text.slice(2);

          // Use outlineLevel instead of heading style to avoid template's automatic numbering

          children.push(new Paragraph({

            outlineLevel: 0, // Level 1 heading (0-indexed)

            spacing: { before: 400, after: 200 },

            children: [new TextRun({ text: headingText, bold: true, color: tplStyles.heading1Color, size: tplStyles.heading1Size, font: tplStyles.heading1Font })]

          }));

          return;

        }

        const stopWords = ["such as", "including", "following"];

        if (/^[A-Z][A-Za-z0-9 ()/-]{3,60}$/.test(text) && text.split(" ").length <= 6 && !stopWords.some(w => text.toLowerCase().includes(w)) && nextLine.startsWith("- ")) {

          children.push(new Paragraph({

            spacing: { before: 250, after: 120 },

            children: [new TextRun({ text, bold: true, color: tplStyles.heading2Color, size: tplStyles.heading2Size, font: tplStyles.heading2Font })]

          }));

          return;

        }

        if (/^[A-Z][A-Za-z0-9 ()/-]{3,80}:$/.test(text)) {

          children.push(new Paragraph({

            spacing: { before: 250, after: 120 },

            children: [new TextRun({ text: text.replace(/:$/, ""), bold: true, color: tplStyles.heading2Color, size: tplStyles.heading2Size, font: tplStyles.heading2Font })]

          }));

          return;

        }

        if (/^[A-Za-z][A-Za-z0-9 ()/-]{1,40}:/.test(text)) {

          children.push(new Paragraph({

            children: splitBoldBeforeColon(text, false, tplStyles.defaultFont)

          }));

          return;

        }

        const labeled = splitNumberedLabel(text);

        if (labeled) {

          children.push(new Paragraph({

            children: [

              new TextRun({

                text: labeled.prefix,

                bold: true,

                color: tplStyles.labelColor,

                font: tplStyles.defaultFont

              }),

              new TextRun({ text: labeled.remainder, font: tplStyles.defaultFont })

            ]

          }));

          return;

        }

        text = normalizeBullet(text);

        if (text.startsWith("- ")) {

          const bulletText = text.slice(2);

          children.push(new Paragraph({

            children: bulletText.includes(":") ? splitBoldBeforeColon(bulletText, true, tplStyles.defaultFont) : [new TextRun({ text: "- " + bulletText, font: tplStyles.defaultFont })]

          }));

          return;

        }

        if (text.includes("|") && text.split("|").length >= 3) {

          insideTable = true;

          currentTableRows.push(text);

          return;

        }

        if (insideTable && text.includes("|")) {

          currentTableRows.push(text);

          return;

        }

        if (insideTable && currentTableRows.length > 0 && !(text.includes("|") && text.split("|").length >= 3)) {

          const table = parseMarkdownTable(currentTableRows);



          if (table) {

            children.push(

              new Paragraph({ text: "" }),

              table,

              new Paragraph({ text: "" })

            );

          } else {

            // fallback so table never disappears

            currentTableRows.forEach(r => {

              children.push(new Paragraph({ text: r }));

            });

          }



          currentTableRows = [];

          insideTable = false;

        }

        children.push(new Paragraph({

          children: [new TextRun({ text, font: tplStyles.defaultFont })]

        }));

      });



      if (insideTable && currentTableRows.length > 0) {

        const table = parseMarkdownTable(currentTableRows);



        if (table) {

          children.push(

            new Paragraph({ text: "" }),

            table,

            new Paragraph({ text: "" })

          );

        } else {

          currentTableRows.forEach(r => {

            children.push(new Paragraph({ text: r }));

          });

        }

      }



      const hasTOC = templateHasTOC(templateDoc);



      // Function to clear existing TOC entries and mark for update

      function clearAndUpdateTOC(templateDoc) {

        const body = templateDoc.getElementsByTagName("w:body")[0];

        if (!body) return;



        // Find all paragraphs in the document

        const paragraphs = body.getElementsByTagName("w:p");

        const tocParagraphsToRemove = [];

        let inTocSection = false;

        let tocFieldDepth = 0;



        // Iterate through paragraphs to find TOC content

        for (let i = 0; i < paragraphs.length; i++) {

          const para = paragraphs[i];



          // Check for TOC field begin

          const fldCharBegins = para.getElementsByTagName("w:fldChar");

          for (let j = 0; j < fldCharBegins.length; j++) {

            const fldType = fldCharBegins[j].getAttribute("w:fldCharType");

            if (fldType === "begin") {

              // Check if this is a TOC field

              const instrTexts = para.getElementsByTagName("w:instrText");

              for (let k = 0; k < instrTexts.length; k++) {

                if (instrTexts[k].textContent && instrTexts[k].textContent.includes("TOC")) {

                  inTocSection = true;

                  tocFieldDepth++;

                }

              }

            } else if (fldType === "end" && inTocSection) {

              tocFieldDepth--;

              if (tocFieldDepth <= 0) {

                inTocSection = false;

                tocFieldDepth = 0;

              }

            }

          }



          if (!inTocSection) {

            const instrTexts = para.getElementsByTagName("w:instrText");

            for (let k = 0; k < instrTexts.length; k++) {

              if (instrTexts[k].textContent && instrTexts[k].textContent.includes("TOC")) {

                inTocSection = true;

                break;

              }

            }

          }





          const hyperlinks = para.getElementsByTagName("w:hyperlink");

          const pStyle = para.getElementsByTagName("w:pStyle");

          let isTocEntry = false;





          for (let j = 0; j < pStyle.length; j++) {

            const styleVal = pStyle[j].getAttribute("w:val");

            if (styleVal && /^TOC\d+$/i.test(styleVal)) {

              isTocEntry = true;

              break;

            }

          }





          for (let j = 0; j < hyperlinks.length; j++) {

            const anchor = hyperlinks[j].getAttribute("w:anchor");

            if (anchor && anchor.startsWith("_Toc")) {

              isTocEntry = true;

              break;

            }

          }





          if (isTocEntry && !para.getElementsByTagName("w:instrText").length) {

            tocParagraphsToRemove.push(para);

          }

        }





        tocParagraphsToRemove.forEach(para => {

          if (para.parentNode) {

            para.parentNode.removeChild(para);

          }

        });





        const instrTexts = templateDoc.getElementsByTagName("w:instrText");

        for (let i = 0; i < instrTexts.length; i++) {

          let txt = instrTexts[i].textContent;

          if (txt && txt.includes("TOC")) {



            instrTexts[i].textContent = txt.replace(/\\o\s*"1-\d+"/, '\\o "1-4"');

          }

        }





        const settings = templateDoc.getElementsByTagName("w:settings");

        if (settings.length === 0) {

        }

      }



      if (hasTOC) {



        clearAndUpdateTOC(templateDoc);

      }



      const docChildren = [];

      if (!hasTOC) {

        docChildren.push(

          new Paragraph({

            text: "Table of Contents",

            heading: HeadingLevel.HEADING_1

          }),

          new TableOfContents("TOC", {

            hyperlink: true,

            headingStyleRange: "1-4"

          }),

          new Paragraph({ pageBreakBefore: true })

        );

      }

      docChildren.push(...children);

      const tempDoc = new Document({

        styles: {

          default: {

            document: {

              run: { font: tplStyles.defaultFont, size: tplStyles.defaultSize }

            }

          }

        },

        sections: [

          {

            children: docChildren

          }

        ]

      });

      const formattedBuffer = await Packer.toBuffer(tempDoc);



      const formattedZip = new PizZip(formattedBuffer);

      const formattedDoc = parser.parseFromString(formattedZip.file("word/document.xml").asText());

      const templateBody = templateDoc.getElementsByTagName("w:body")[0];

      const formattedBody = formattedDoc.getElementsByTagName("w:body")[0];



      // ─── Section-to-Section Mapping Logic ───

      // Helper: extract text content from a w:p element

      function getParagraphText(paraNode) {

        const runs = paraNode.getElementsByTagName("w:t");

        let text = "";

        for (let i = 0; i < runs.length; i++) {

          text += runs[i].textContent || "";

        }

        return text.trim();

      }



      // Helper: determine heading level of a paragraph (0 = not a heading)

      function getHeadingLevel(paraNode) {

        const pPr = paraNode.getElementsByTagName("w:pPr");

        if (pPr.length === 0) return 0;



        const pStyle = pPr[0].getElementsByTagName("w:pStyle");

        if (pStyle.length > 0) {

          const styleVal = pStyle[0].getAttribute("w:val") || "";

          const headingMatch = styleVal.match(/^Heading(\d)$/i);

          if (headingMatch) return parseInt(headingMatch[1], 10);

          // Skip TOC styles

          if (/^TOC\d$/i.test(styleVal)) return 0;

        }



        const outlineLvl = pPr[0].getElementsByTagName("w:outlineLvl");

        if (outlineLvl.length > 0) {

          const lvl = parseInt(outlineLvl[0].getAttribute("w:val"), 10);

          if (lvl >= 0 && lvl <= 5) return lvl + 1;

        }

        return 0;

      }



      // Helper: normalize heading text for comparison (case-insensitive, remove numbering prefixes, extra whitespace)

      function normalizeHeading(text) {

        return text

          .replace(/^\d+(\.\d+)*[\.\)]\s*/g, "") // Remove leading "1. ", "1.2. ", "1.2.1 " etc.

          .replace(/^\d+(\.\d+)*\s+/g, "")       // Remove leading "1.2.1 " without period

          .replace(/^SECTION\s+\d+\s*[-—:.]?\s*/i, "") // Remove "SECTION 1 - " prefix

          .replace(/[━─—\-:]/g, " ")           // Replace dashes/colons with space

          .replace(/\s+/g, " ")                // Collapse whitespace

          .trim()

          .toLowerCase();

      }



      // Parse formatted AI content into sections: { normalizedHeading -> [nodes] }

      function parseFormattedSections(bodyNode) {

        const allNodes = Array.from(bodyNode.childNodes).filter(

          n => n.nodeType === 1 && n.nodeName !== "w:sectPr"

        );



        const sections = [];

        let currentSection = null;



        allNodes.forEach(node => {

          if (node.nodeName === "w:p") {

            const level = getHeadingLevel(node);

            const text = getParagraphText(node);



            // Treat all heading levels (1-6) as section boundaries

            if (level >= 1 && level <= 6 && text) {

              // Start a new section

              currentSection = {

                headingText: text,

                normalizedKey: normalizeHeading(text),

                level: level,

                nodes: [node]  // Include the heading node itself

              };

              sections.push(currentSection);

              return;

            }

          }



          // Add to current section or create a preamble section

          if (currentSection) {

            currentSection.nodes.push(node);

          } else {

            // Nodes before the first heading - create a preamble section

            if (!sections.length || sections[0].normalizedKey !== "__preamble__") {

              currentSection = {

                headingText: "__preamble__",

                normalizedKey: "__preamble__",

                level: 0,

                nodes: [node]

              };

              sections.unshift(currentSection);

            } else {

              sections[0].nodes.push(node);

            }

          }

        });



        return sections;

      }



      // Parse template body to find heading positions (all levels 1-6)

      function parseTemplateHeadings(bodyNode) {

        // ── Helpers used both here and by the "inferred sub-heading" pass ──

        // Check if a paragraph node is inside a table (w:tbl ancestor).

        function isInsideTablePara(node) {

          let p = node.parentNode;

          while (p) {

            if (p.nodeName === "w:tbl") return true;

            p = p.parentNode;

          }

          return false;

        }

        // Check whether all runs in a paragraph are bold. Returns true only if

        // there is at least one w:r AND every run's rPr has w:b (bold marker).

        function isAllRunsBold(paraNode) {

          const runs = paraNode.getElementsByTagName("w:r");

          if (runs.length === 0) return false;

          let anyTextRun = false;

          for (let r = 0; r < runs.length; r++) {

            const texts = runs[r].getElementsByTagName("w:t");

            if (texts.length === 0) continue; // ignore non-text runs (breaks, fldChar, etc.)

            anyTextRun = true;

            const rPrs = runs[r].getElementsByTagName("w:rPr");

            let boldOnThisRun = false;

            if (rPrs.length > 0) {

              // Look at direct rPr children only (avoid rPr inside nested elements)

              for (let x = 0; x < rPrs.length; x++) {

                if (rPrs[x].parentNode !== runs[r]) continue;

                const bs = rPrs[x].getElementsByTagName("w:b");

                for (let y = 0; y < bs.length; y++) {

                  if (bs[y].parentNode !== rPrs[x]) continue;

                  // <w:b/> present means bold ON unless w:val="0"/"false"

                  const v = bs[y].getAttribute("w:val");

                  if (!v || v === "1" || v.toLowerCase() === "true") {

                    boldOnThisRun = true;

                  }

                }

              }

            }

            if (!boldOnThisRun) return false;

          }

          return anyTextRun;

        }



        // First try direct children of w:body

        const directChildren = Array.from(bodyNode.childNodes).filter(n => n.nodeType === 1);

        const headings = [];



        directChildren.forEach((node, index) => {

          if (node.nodeName === "w:p") {

            const level = getHeadingLevel(node);

            const text = getParagraphText(node);



            // Include all heading levels for matching

            if (level >= 1 && level <= 6 && text) {

              headings.push({

                node: node,

                text: text,

                normalizedKey: normalizeHeading(text),

                level: level,

                index: index

              });

            }

          }

        });



        // If no headings found as direct children, search ALL w:p elements

        // (headings may be inside w:sdt or other wrapper elements)

        if (headings.length === 0) {

          console.log("[generateDocument] No direct-child headings found, searching all paragraphs...");

          const allParagraphs = bodyNode.getElementsByTagName("w:p");

          for (let i = 0; i < allParagraphs.length; i++) {

            const node = allParagraphs[i];

            const level = getHeadingLevel(node);

            if (level === 0) continue;

            const text = getParagraphText(node);

            if (!text) continue;

            // Skip TOC entries

            const pPr = node.getElementsByTagName("w:pPr");

            if (pPr.length > 0) {

              const pStyle = pPr[0].getElementsByTagName("w:pStyle");

              if (pStyle.length > 0) {

                const styleVal = pStyle[0].getAttribute("w:val") || "";

                if (/^TOC\d/i.test(styleVal)) continue;

              }

            }

            headings.push({

              node: node,

              text: text,

              normalizedKey: normalizeHeading(text),

              level: level,

              index: i

            });

          }

        }





        // Fallback: if no headings found via styles, detect from paragraph text patterns

        if (headings.length === 0) {

          console.log("[generateDocument] No styled template headings found, using text-based detection");

          const allParas = bodyNode.getElementsByTagName("w:p");

          for (let pi = 0; pi < allParas.length; pi++) {

            const pNode = allParas[pi];

            const pText = getParagraphText(pNode).trim();

            if (!pText || pText.length > 100) continue;

            // Skip TOC styled paragraphs

            const pPr = pNode.getElementsByTagName("w:pPr");

            if (pPr.length > 0) {

              const pStyle = pPr[0].getElementsByTagName("w:pStyle");

              if (pStyle.length > 0) {

                const sv = pStyle[0].getAttribute("w:val") || "";

                if (/^TOC/i.test(sv)) continue;

              }

            }

            // Detect numbered headings: "1. Conceptual Design" or "1.1 Brief description"

            const numMatch = pText.match(/^(\d+(\.\d+)*)[\.\)]?\s+(.+)$/);

            if (numMatch) {

              const dots = (numMatch[1].match(/\./g) || []).length;

              headings.push({ node: pNode, text: pText, normalizedKey: normalizeHeading(pText), level: dots + 1, index: pi });

              continue;

            }

            // Detect ALL CAPS short lines as headings

            if (pText.length > 3 && pText.length < 60 && pText === pText.toUpperCase() && /[A-Z]/.test(pText)) {

              headings.push({ node: pNode, text: pText, normalizedKey: normalizeHeading(pText), level: 1, index: pi });

              continue;

            }

          }

        }



        // ─── Supplementary pass: infer sub-headings from formatting/structure ───

        //

        // Many .docx templates (FSDs in particular) use plain Normal paragraphs

        // as sub-headings instead of real Heading2/3 styles. These may be:

        //   (a) bold-only Normal paragraphs, OR

        //   (b) plain Normal paragraphs that just look like section titles

        //       (short, Title Case, no ending punctuation) — this is the

        //       common case in the Rolex/Signify FSD templates.

        //

        // Without detecting them, the section-matching logic treats the whole

        // document as one giant section, marks it "not placeholder" (too much

        // text) and appends all AI content at the end instead of replacing

        // placeholders section-by-section.

        //

        // Because a plain-Normal detector is inherently fuzzy, we apply it

        // only to paragraphs that look strongly heading-like AND are

        // followed by non-heading content (i.e. they actually introduce a

        // section rather than being just a stray label).



        // Helper: does this paragraph contain a hyperlink anchor to a TOC

        // bookmark (e.g. w:hyperlink @w:anchor starting with _Toc)?

        // Such paragraphs are ToC entries, not section headings.

        function isTocHyperlinkPara(node) {

          const hyperlinks = node.getElementsByTagName("w:hyperlink");

          for (let i = 0; i < hyperlinks.length; i++) {

            const anchor = hyperlinks[i].getAttribute("w:anchor") || "";

            if (anchor.startsWith("_Toc")) return true;

          }

          return false;

        }



        // Helper: does this paragraph "look" like a section heading?

        // - short (< 90 chars)

        // - between 1 and 10 words

        // - doesn't end with .,;: or ? or !

        // - starts with a capital letter or a digit (numbered heading)

        // - doesn't contain a colon inside (labels like "Description:" are

        //   inline field labels, not section headings)

        function looksLikeHeadingText(text) {

          if (!text) return false;

          const t = text.trim();

          if (t.length < 2 || t.length > 90) return false;

          if (/[.,;?!:]$/.test(t)) {

            // Allow trailing colon on a very short label like "Overall Enhancement Requirement:"

            if (!/^[A-Z][A-Za-z0-9 ()/&\-–—]{2,80}:$/.test(t)) return false;

          }

          if (t.includes(":") && !/:$/.test(t)) return false;

          const words = t.split(/\s+/);

          if (words.length < 1 || words.length > 10) return false;

          if (!/^[A-Z0-9]/.test(t)) return false;

          // Reject sentences that contain typical verb-connector words which

          // rarely appear in section titles.

          if (/\b(is|are|was|were|will|would|should|has|have|had|does|did|the|of|and|or)\b/i.test(t)) {

            // BUT template heading titles often contain "of"/"and" (e.g.

            // "Security and Roles", "Overview of Requirement"). Allow when

            // the paragraph is Title Case-ish (most words capitalised).

            const contentWords = words.filter(w => w.length > 2);

            if (contentWords.length === 0) return false;

            const capWords = contentWords.filter(w => /^[A-Z]/.test(w));

            if (capWords.length / contentWords.length < 0.5) return false;

          }

          return true;

        }



        {

          const existingNodes = new Set(headings.map(h => h.node));

          const seenNormalizedKeys = new Set(headings.map(h => h.normalizedKey));



          // First find the position of the last styled heading so we don't

          // accidentally infer headings inside a wrapped ToC block at the top.

          directChildren.forEach((node, index) => {

            if (node.nodeName !== "w:p") return;

            if (existingNodes.has(node)) return;

            if (isInsideTablePara(node)) return;

            if (isTocHyperlinkPara(node)) return;



            const text = getParagraphText(node);

            if (!text) return;



            // Skip TOC-styled paragraphs and instrText field paragraphs

            const pPr = node.getElementsByTagName("w:pPr");

            if (pPr.length > 0) {

              const pStyle = pPr[0].getElementsByTagName("w:pStyle");

              if (pStyle.length > 0) {

                const styleVal = pStyle[0].getAttribute("w:val") || "";

                if (/^TOC\d/i.test(styleVal)) return;

              }

            }

            const instrTexts = node.getElementsByTagName("w:instrText");

            if (instrTexts.length > 0) return;



            const bold = isAllRunsBold(node);

            const looksLikeHeading = looksLikeHeadingText(text);



            // Accept as inferred heading if:

            //   - it's bold (strong signal), OR

            //   - it looks like a title AND is not a full sentence

            if (!bold && !looksLikeHeading) return;



            const normalizedKey = normalizeHeading(text);

            // Skip if we already have a heading with this normalized key

            // (avoids duplicating ToC-style entries as body headings).

            if (seenNormalizedKeys.has(normalizedKey)) return;



            headings.push({

              node: node,

              text: text,

              normalizedKey: normalizedKey,

              level: 2,

              index: index,

              inferred: true

            });

            seenNormalizedKeys.add(normalizedKey);

          });



          // Sort headings by their position in the document, so the section

          // slicing that follows uses the correct order.

          headings.sort((a, b) => (a.index || 0) - (b.index || 0));

          if (headings.some(h => h.inferred)) {

            const inferredCount = headings.filter(h => h.inferred).length;

            console.log("[generateDocument] Inferred", inferredCount, "sub-headings in template");

          }

        }



        return headings;

      }



      // Get content nodes between two heading positions in template

      function getContentBetweenHeadings(bodyNode, headingNode, nextHeadingNode) {

        const directChildren = Array.from(bodyNode.childNodes).filter(n => n.nodeType === 1);

        const contentNodes = [];

        let collecting = false;



        // Find the direct-child ancestor of headingNode (handles w:sdt wrappers)

        function findDirectChildAncestor(node) {

          if (!node) return null;

          if (node.parentNode === bodyNode) return node;

          // Walk up to find the ancestor that is a direct child of body

          let current = node;

          while (current && current.parentNode !== bodyNode) {

            current = current.parentNode;

          }

          return current;

        }



        const headingAncestor = findDirectChildAncestor(headingNode);

        const nextHeadingAncestor = nextHeadingNode ? findDirectChildAncestor(nextHeadingNode) : null;



        for (const node of directChildren) {

          // Check if this node IS or CONTAINS the heading

          if (node === headingNode || node === headingAncestor) {

            collecting = true;

            continue; // Skip the heading/wrapper itself

          }

          // Check if this node IS or CONTAINS the next heading

          if (nextHeadingNode && (node === nextHeadingNode || node === nextHeadingAncestor)) {

            break; // Stop at next heading

          }

          if (node.nodeName === "w:sectPr") {

            break; // Stop at sectPr

          }

          if (collecting) {

            contentNodes.push(node);

          }

        }



        return contentNodes;

      }



      // Check if content between headings is "placeholder" (empty or very short generic text)

      function isPlaceholderContent(contentNodes) {

        if (contentNodes.length === 0) return true;



        let totalText = "";

        contentNodes.forEach(node => {

          if (node.nodeName === "w:p") {

            totalText += getParagraphText(node) + " ";

          }

        });

        totalText = totalText.trim();



        // Consider it a placeholder if text is empty or very short (< 50 chars)

        // or if it contains typical placeholder patterns

        if (!totalText) return true;

        if (totalText.length < 50) return true;

        if (/^\[.*\]$/.test(totalText)) return true; // [Placeholder text]

        if (/^<.*>$/.test(totalText)) return true;   // <Placeholder text>

        if (/lorem ipsum/i.test(totalText)) return true;

        if (/insert\s+(content|text|here)/i.test(totalText)) return true;

        if (/to\s+be\s+(completed|filled|added|updated|mentioned|provided)/i.test(totalText)) return true;

        if (/^\(.*\)$/.test(totalText)) return true;

        if (/^\(here|in this section|a brief|detailed|provide|mention/i.test(totalText)) return true;

        if (/not\s+available/i.test(totalText) && totalText.length < 100) return true;

        if (/will\s+be\s+(mentioned|updated|provided|added)/i.test(totalText)) return true;

        if (/please\s+(refer|see|provide)/i.test(totalText) && totalText.length < 150) return true;





        // Detect structured placeholder: sections with mostly short label-like paragraphs

        // (e.g., "Assumptions:" "Business" "Technical" "Dependencies:" etc.)

        let shortParaCount = 0;

        let totalParaCount = 0;

        contentNodes.forEach(node => {

          if (node.nodeName === "w:p") {

            const pText = getParagraphText(node).trim();

            if (pText) {

              totalParaCount++;

              if (pText.length < 25) shortParaCount++;

            }

          }

        });

        if (totalParaCount >= 3 && (shortParaCount / totalParaCount) >= 0.7) return true;



        return false;

      }



      // ─── Table Population Logic ───

      // Extract text from a table cell (w:tc)

      function getTableCellText(tcNode) {

        const texts = tcNode.getElementsByTagName("w:t");

        let cellText = "";

        for (let i = 0; i < texts.length; i++) {

          cellText += texts[i].textContent || "";

        }

        return cellText.trim();

      }



      // Extract header row texts from a w:tbl element

      function getTableHeaders(tblNode) {

        const rows = tblNode.getElementsByTagName("w:tr");

        if (rows.length === 0) return [];

        const firstRow = rows[0];

        const cells = firstRow.getElementsByTagName("w:tc");

        const headers = [];

        for (let i = 0; i < cells.length; i++) {

          headers.push(getTableCellText(cells[i]).toLowerCase());

        }

        return headers;

      }



      // Check if two sets of headers are compatible (similar text, flexible column count)

      function headersMatch(tplHeaders, aiHeaders) {

        if (tplHeaders.length === 0 || aiHeaders.length === 0) return false;

        // Allow slight column count differences (±1 column)

        if (Math.abs(tplHeaders.length - aiHeaders.length) > 1) return false;



        const minLen = Math.min(tplHeaders.length, aiHeaders.length);

        let matchCount = 0;

        for (let i = 0; i < minLen; i++) {

          const tH = tplHeaders[i].replace(/[^a-z0-9]/g, "");

          const aH = aiHeaders[i].replace(/[^a-z0-9]/g, "");

          if (tH && aH && (tH.includes(aH) || aH.includes(tH) || tH === aH)) {

            matchCount++;

          }

        }

        // At least 50% of the shorter header set must match

        return matchCount >= Math.ceil(minLen * 0.5);

      }



      // Extract data rows from AI table (skip header row, return array of arrays)

      function getAITableDataRows(tblNode) {

        const rows = tblNode.getElementsByTagName("w:tr");

        const dataRows = [];

        for (let i = 1; i < rows.length; i++) { // Skip header row (index 0)

          const cells = rows[i].getElementsByTagName("w:tc");

          const rowData = [];

          for (let j = 0; j < cells.length; j++) {

            rowData.push(getTableCellText(cells[j]));

          }

          dataRows.push(rowData);

        }

        return dataRows;

      }



      // Check if a template table row is empty (all cells blank)

      function isEmptyTableRow(trNode) {

        const cells = trNode.getElementsByTagName("w:tc");

        for (let i = 0; i < cells.length; i++) {

          if (getTableCellText(cells[i]).trim()) return false;

        }

        return true;

      }



      // Populate template table with AI data rows

      // Returns true if population was successful

      function populateTemplateTable(templateTbl, aiDataRows, tplDoc) {

        const rows = templateTbl.getElementsByTagName("w:tr");

        if (rows.length < 2) return false; // Need at least header + 1 data row template



        // Find the first empty data row to use as a template for new rows

        let templateRow = null;

        const emptyRows = [];

        for (let i = 1; i < rows.length; i++) {

          if (isEmptyTableRow(rows[i])) {

            if (!templateRow) templateRow = rows[i];

            emptyRows.push(rows[i]);

          }

        }



        // If no empty template row found, use the last row as template

        if (!templateRow && rows.length > 1) {

          templateRow = rows[rows.length - 1];

        }

        if (!templateRow) return false;



        // Remove existing empty rows (we'll add populated ones)

        emptyRows.forEach(row => {

          if (row.parentNode) {

            row.parentNode.removeChild(row);

          }

        });



        // For each AI data row, clone the template row and populate cells

        aiDataRows.forEach(rowData => {

          const newRow = templateRow.cloneNode(true);

          const cells = newRow.getElementsByTagName("w:tc");



          for (let j = 0; j < cells.length && j < rowData.length; j++) {

            // Clear existing text in the cell

            const paragraphs = cells[j].getElementsByTagName("w:p");

            if (paragraphs.length > 0) {

              const para = paragraphs[0];

              // Remove all existing runs

              const existingRuns = para.getElementsByTagName("w:r");

              const runsToRemove = [];

              for (let r = 0; r < existingRuns.length; r++) {

                runsToRemove.push(existingRuns[r]);

              }

              runsToRemove.forEach(run => {

                if (run.parentNode === para) para.removeChild(run);

              });



              // Create a new run with the AI data text

              const newRun = tplDoc.createElement("w:r");

              const newText = tplDoc.createElement("w:t");

              newText.setAttribute("xml:space", "preserve");

              newText.textContent = rowData[j] || "";

              newRun.appendChild(newText);

              para.appendChild(newRun);

            }

          }



          // Insert the new row before sectPr or at end of table

          templateTbl.appendChild(newRow);

        });



        return true;

      }



      // Try to populate template tables with AI table data

      // Returns modified AI content nodes (with matched tables removed)

      function handleTablePopulation(existingContent, aiContentNodes, tplDoc) {

        // Find tables in template content

        const templateTables = [];

        existingContent.forEach(node => {

          if (node.nodeName === "w:tbl") {

            templateTables.push(node);

          }

        });



        console.log("[generateDocument] handleTablePopulation: existingContent nodes:", existingContent.length, "template tables found:", templateTables.length);



        if (templateTables.length === 0) {

          console.log("[generateDocument] handleTablePopulation: No template tables in existingContent, skipping");

          return aiContentNodes;

        }



        // Find tables in AI content

        const aiTables = [];

        const aiTableIndices = [];

        aiContentNodes.forEach((node, idx) => {

          if (node.nodeName === "w:tbl") {

            aiTables.push({ node, idx });

            aiTableIndices.push(idx);

          }

        });



        console.log("[generateDocument] handleTablePopulation: AI tables found:", aiTables.length);



        if (aiTables.length === 0) return aiContentNodes; // No AI tables



        // Log headers for debugging

        templateTables.forEach((tbl, i) => {

          const headers = getTableHeaders(tbl);

          const rows = tbl.getElementsByTagName("w:tr");

          let emptyRowCount = 0;

          for (let ri = 1; ri < rows.length; ri++) {

            if (isEmptyTableRow(rows[ri])) emptyRowCount++;

          }

          console.log(`[generateDocument] Template table ${i}: headers=[${headers.join(", ")}] totalRows=${rows.length} emptyRows=${emptyRowCount}`);

        });

        aiTables.forEach((aiTbl, i) => {

          const headers = getTableHeaders(aiTbl.node);

          console.log(`[generateDocument] AI table ${i}: headers=[${headers.join(", ")}]`);

        });



        const populatedAITableIndices = new Set();

        const populatedTemplateTableIndices = new Set();



        // Helper: check if a template table is empty (all data rows blank)

        function isTemplateTableEmpty(tblNode) {

          const rows = tblNode.getElementsByTagName("w:tr");

          for (let i = 1; i < rows.length; i++) {

            if (!isEmptyTableRow(rows[i])) return false;

          }

          return true;

        }



        // First pass: prioritize populating EMPTY template tables

        templateTables.forEach((tplTable, tplIdx) => {

          if (!isTemplateTableEmpty(tplTable)) return; // Skip non-empty tables in first pass

          const tplHeaders = getTableHeaders(tplTable);

          if (tplHeaders.length === 0) return;



          for (const aiTable of aiTables) {

            if (populatedAITableIndices.has(aiTable.idx)) continue;



            const aiHeaders = getTableHeaders(aiTable.node);

            if (headersMatch(tplHeaders, aiHeaders)) {

              const aiDataRows = getAITableDataRows(aiTable.node);

              if (aiDataRows.length > 0) {

                const success = populateTemplateTable(tplTable, aiDataRows, tplDoc);

                if (success) {

                  populatedAITableIndices.add(aiTable.idx);

                  populatedTemplateTableIndices.add(tplIdx);

                  console.log("[generateDocument] Populated empty template table with", aiDataRows.length, "rows. Headers:", tplHeaders.join(", "));

                }

              }

              break;

            }

          }

        });



        // Second pass: try remaining AI tables against non-empty template tables (append mode)

        templateTables.forEach((tplTable, tplIdx) => {

          if (populatedTemplateTableIndices.has(tplIdx)) return; // Already populated

          if (isTemplateTableEmpty(tplTable)) return; // Already tried in first pass

          const tplHeaders = getTableHeaders(tplTable);

          if (tplHeaders.length === 0) return;



          for (const aiTable of aiTables) {

            if (populatedAITableIndices.has(aiTable.idx)) continue;



            const aiHeaders = getTableHeaders(aiTable.node);

            if (headersMatch(tplHeaders, aiHeaders)) {

              const aiDataRows = getAITableDataRows(aiTable.node);

              if (aiDataRows.length > 0) {

                const success = populateTemplateTable(tplTable, aiDataRows, tplDoc);

                if (success) {

                  populatedAITableIndices.add(aiTable.idx);

                  populatedTemplateTableIndices.add(tplIdx);

                  console.log("[generateDocument] Populated non-empty template table with", aiDataRows.length, "rows. Headers:", tplHeaders.join(", "));

                }

              }

              break;

            }

          }

        });



        // Return AI nodes without the tables that were already populated

        if (populatedAITableIndices.size === 0) return aiContentNodes;



        return aiContentNodes.filter((_, idx) => !populatedAITableIndices.has(idx));

      }



      // ─── Text-based section detection fallback ───

      // When AI content doesn't use proper heading styles, detect sections by matching paragraph text

      function parseFormattedSectionsByText(bodyNode, templateHeadingKeys) {

        const allNodes = Array.from(bodyNode.childNodes).filter(

          n => n.nodeType === 1 && n.nodeName !== "w:sectPr"

        );



        const sections = [];

        let currentSection = null;



        allNodes.forEach(node => {

          if (node.nodeName === "w:p") {

            const text = getParagraphText(node);

            if (text) {

              const normalizedText = normalizeHeading(text);

              // Check if this paragraph's text matches any template heading

              if (normalizedText && templateHeadingKeys.has(normalizedText)) {

                currentSection = {

                  headingText: text,

                  normalizedKey: normalizedText,

                  level: 1,

                  nodes: [node]

                };

                sections.push(currentSection);

                return;

              }

              // Also check if any template heading is a substring of this text or vice versa

              for (const tplKey of templateHeadingKeys) {

                if (tplKey.length > 3 && (normalizedText.includes(tplKey) || tplKey.includes(normalizedText)) && normalizedText.length < 100) {

                  currentSection = {

                    headingText: text,

                    normalizedKey: tplKey, // Use the template key for matching

                    level: 1,

                    nodes: [node]

                  };

                  sections.push(currentSection);

                  return;

                }

              }

            }

          }



          if (currentSection) {

            currentSection.nodes.push(node);

          } else {

            if (!sections.length || sections[0].normalizedKey !== "__preamble__") {

              currentSection = {

                headingText: "__preamble__",

                normalizedKey: "__preamble__",

                level: 0,

                nodes: [node]

              };

              sections.unshift(currentSection);

            } else {

              sections[0].nodes.push(node);

            }

          }

        });



        return sections;

      }



      // ─── Global Table Population Pass ───

      // Scan ALL tables in the entire template body and try to match with AI tables

      // This handles cases where tables might not be found within specific section boundaries

      (function globalTablePopulation() {

        const allTemplateTables = templateBody.getElementsByTagName("w:tbl");

        const allAITables = formattedBody.getElementsByTagName("w:tbl");



        if (allTemplateTables.length === 0 || allAITables.length === 0) return;



        console.log(`[generateDocument] Global table pass: ${allTemplateTables.length} template tables, ${allAITables.length} AI tables`);



        const globalPopulatedAI = new Set();



        // Find empty template tables and try to match with AI tables

        for (let ti = 0; ti < allTemplateTables.length; ti++) {

          const tplTable = allTemplateTables[ti];

          const tplHeaders = getTableHeaders(tplTable);

          if (tplHeaders.length === 0) continue;



          // Check if this template table is empty

          const tplRows = tplTable.getElementsByTagName("w:tr");

          let isEmpty = true;

          for (let ri = 1; ri < tplRows.length; ri++) {

            if (!isEmptyTableRow(tplRows[ri])) { isEmpty = false; break; }

          }

          if (!isEmpty) continue; // Skip non-empty template tables



          // Try to match with an AI table

          for (let ai = 0; ai < allAITables.length; ai++) {

            if (globalPopulatedAI.has(ai)) continue;

            const aiTable = allAITables[ai];

            const aiHeaders = getTableHeaders(aiTable);



            if (headersMatch(tplHeaders, aiHeaders)) {

              const aiDataRows = getAITableDataRows(aiTable);

              if (aiDataRows.length > 0) {

                const success = populateTemplateTable(tplTable, aiDataRows, templateDoc);

                if (success) {

                  globalPopulatedAI.add(ai);

                  console.log(`[generateDocument] GLOBAL: Populated template table ${ti} with ${aiDataRows.length} AI rows. Headers: [${tplHeaders.join(", ")}]`);



                  // Mark this AI table for removal from formatted content

                  // by adding a data attribute we can check later

                  aiTable.setAttribute("data-populated", "true");

                }

              }

              break; // One match per template table

            }

          }

        }

      })();



      // ─── Perform Section Mapping ───

      const templateHeadings = parseTemplateHeadings(templateBody);

      let aiSections = parseFormattedSections(formattedBody);



      // If formal heading detection found <= 1 section, try text-based detection

      if (aiSections.length <= 1 && templateHeadings.length > 0) {

        console.log("[generateDocument] No formal AI headings detected, trying text-based section matching...");

        const templateHeadingKeys = new Set(templateHeadings.map(h => h.normalizedKey));

        const textBasedSections = parseFormattedSectionsByText(formattedBody, templateHeadingKeys);

        if (textBasedSections.length > 1) {

          aiSections = textBasedSections;

          console.log("[generateDocument] Text-based detection found", aiSections.length, "sections");

        }

      }



      console.log("[generateDocument] Template headings found:", templateHeadings.map(h => `[L${h.level}] ${h.text}`));

      console.log("[generateDocument] AI sections found:", aiSections.map(s => `[L${s.level}] ${s.headingText}`));



      // Build a map from normalized AI heading text -> AI sections (array to handle duplicates)

      const aiSectionMap = new Map();

      aiSections.forEach(section => {

        if (section.normalizedKey !== "__preamble__") {

          if (!aiSectionMap.has(section.normalizedKey)) {

            aiSectionMap.set(section.normalizedKey, []);

          }

          aiSectionMap.get(section.normalizedKey).push(section);

        }

      });



      // Track how many times each normalized key has been consumed (for duplicate handling)

      const aiSectionConsumedCount = new Map();



      // Try to match template headings to AI sections

      let matchedSections = new Set();

      let sectionMappingUsed = false;



      if (templateHeadings.length > 0 && aiSections.length > 1) {

        // Attempt fuzzy matching for each template heading

        const matches = [];



        templateHeadings.forEach((tplHeading, idx) => {

          let matchedAI = null;



          // Direct match (consume in order for duplicate keys)

          const directMatches = aiSectionMap.get(tplHeading.normalizedKey);

          if (directMatches && directMatches.length > 0) {

            const consumedCount = aiSectionConsumedCount.get(tplHeading.normalizedKey) || 0;

            if (consumedCount < directMatches.length) {

              matchedAI = directMatches[consumedCount];

              aiSectionConsumedCount.set(tplHeading.normalizedKey, consumedCount + 1);

            }

          }



          // If no direct match, try substring matching

          if (!matchedAI) {

            for (const [key, sections] of aiSectionMap.entries()) {

              if (key.includes(tplHeading.normalizedKey) || tplHeading.normalizedKey.includes(key)) {

                const consumedCount = aiSectionConsumedCount.get(key) || 0;

                if (consumedCount < sections.length) {

                  matchedAI = sections[consumedCount];

                  aiSectionConsumedCount.set(key, consumedCount + 1);

                  break;

                }

              }

            }

          }



          // Try word overlap matching (at least 60% of words match)

          if (!matchedAI) {

            const tplWords = tplHeading.normalizedKey.split(/\s+/).filter(w => w.length > 2);

            let bestMatch = null;

            let bestOverlap = 0;

            let bestKey = null;



            for (const [key, sections] of aiSectionMap.entries()) {

              const consumedCount = aiSectionConsumedCount.get(key) || 0;

              if (consumedCount >= sections.length) continue;

              const aiWords = key.split(/\s+/).filter(w => w.length > 2);

              const overlap = tplWords.filter(w => aiWords.includes(w)).length;

              const overlapRatio = tplWords.length > 0 ? overlap / tplWords.length : 0;



              if (overlapRatio > 0.6 && overlap > bestOverlap) {

                bestOverlap = overlap;

                bestMatch = sections[consumedCount];

                bestKey = key;

              }

            }



            if (bestMatch) {

              matchedAI = bestMatch;

              const consumedCount = aiSectionConsumedCount.get(bestKey) || 0;

              aiSectionConsumedCount.set(bestKey, consumedCount + 1);

            }

          }



          if (matchedAI) {

            matches.push({ templateHeading: tplHeading, aiSection: matchedAI });

            matchedSections.add(matchedAI);

          }

        });



        console.log("[generateDocument] Section matches:", matches.map(m => `"${m.templateHeading.text}" → "${m.aiSection.headingText}"`));



        // If we matched at least some sections, use section-mapping mode

        if (matches.length > 0) {

          sectionMappingUsed = true;



          // For each matched section: remove placeholder content and insert AI content

          matches.forEach(({ templateHeading, aiSection }) => {

            const tplIdx = templateHeadings.indexOf(templateHeading);

            const nextHeading = templateHeadings[tplIdx + 1];

            const nextHeadingNode = nextHeading ? nextHeading.node : null;



            // Get existing content between this heading and the next

            const existingContent = getContentBetweenHeadings(templateBody, templateHeading.node, nextHeadingNode);



            // Get AI content nodes (skip heading)

            let aiContentNodes = aiSection.nodes.slice(1);



            // If this AI section has no content (only heading, e.g., parent section immediately

            // followed by sub-heading), skip it — don't remove existing template content

            if (aiContentNodes.length === 0) {

              return;

            }



            // Try table population first — populate template tables with AI data

            aiContentNodes = handleTablePopulation(existingContent, aiContentNodes, templateDoc);



            // Decide whether to REPLACE or APPEND based on existing content

            // If existing content is just placeholder text, replace it entirely

            // If existing content has real/substantive data, keep it and append AI content after

            const existingIsPlaceholder = isPlaceholderContent(existingContent);



            if (existingIsPlaceholder) {

              // Replace mode: remove placeholder content

              existingContent.forEach(node => {

                if (node.parentNode === templateBody) {

                  if (node.nodeName === "w:tbl") {

                    const rows = node.getElementsByTagName("w:tr");

                    let hasData = false;

                    for (let ri = 1; ri < rows.length; ri++) {

                      if (!isEmptyTableRow(rows[ri])) { hasData = true; break; }

                    }

                    if (hasData) return; // Keep populated table

                  }

                  templateBody.removeChild(node);

                }

              });

            }

            // else: Keep existing content as-is (append mode - AI content will be inserted after it)

            else {

              // Append mode: keep existing content BUT remove empty template tables

              // (since AI will provide its own populated tables)

              existingContent.forEach(node => {

                if (node.parentNode === templateBody && node.nodeName === "w:tbl") {

                  const rows = node.getElementsByTagName("w:tr");

                  let hasData = false;

                  for (let ri = 1; ri < rows.length; ri++) {

                    if (!isEmptyTableRow(rows[ri])) { hasData = true; break; }

                  }

                  if (!hasData) {

                    templateBody.removeChild(node);

                  }

                }

              });

            }



            // Insert remaining AI section content (tables already populated are removed from list)

            const insertBefore = nextHeadingNode || templateBody.getElementsByTagName("w:sectPr")[0] || null;



            aiContentNodes.forEach(node => {

              const imported = templateDoc.importNode(node, true);

              if (insertBefore && insertBefore.parentNode === templateBody) {

                templateBody.insertBefore(imported, insertBefore);

              } else {

                // Insert before the last sectPr

                const lastSectPr = templateBody.getElementsByTagName("w:sectPr");

                if (lastSectPr.length > 0) {

                  const lastSect = lastSectPr[lastSectPr.length - 1];

                  if (lastSect.parentNode === templateBody) {

                    templateBody.insertBefore(imported, lastSect);

                  } else {

                    templateBody.appendChild(imported);

                  }

                } else {

                  templateBody.appendChild(imported);

                }

              }

            });

          });



          // Append any unmatched AI sections at the end (before sectPr)

          const unmatchedSections = aiSections.filter(s =>

            s.normalizedKey !== "__preamble__" && !matchedSections.has(s)

          );



          if (unmatchedSections.length > 0) {

            console.log("[generateDocument] Appending unmatched sections:", unmatchedSections.map(s => s.headingText));

            const lastSectPr = templateBody.getElementsByTagName("w:sectPr");

            const insertRef = lastSectPr.length > 0 ? lastSectPr[lastSectPr.length - 1] : null;



            unmatchedSections.forEach(section => {

              section.nodes.forEach(node => {

                const imported = templateDoc.importNode(node, true);

                if (insertRef && insertRef.parentNode === templateBody) {

                  templateBody.insertBefore(imported, insertRef);

                } else {

                  templateBody.appendChild(imported);

                }

              });

            });

          }

        }

      }



      // ─── Fallback: If no section mapping was possible, use original append logic ───

      if (!sectionMappingUsed) {

        console.log("[generateDocument] No section mapping possible, using append mode");



        const allSectPr = templateBody.getElementsByTagName("w:sectPr");

        let sectPr = null;



        if (allSectPr.length >= 4) {

          const thirdToLastSectPr = allSectPr[allSectPr.length - 3];

          let parentOfSectPr = thirdToLastSectPr.parentNode;

          while (parentOfSectPr && parentOfSectPr.parentNode !== templateBody) {

            parentOfSectPr = parentOfSectPr.parentNode;

          }

          if (parentOfSectPr && parentOfSectPr.nextSibling) {

            sectPr = parentOfSectPr.nextSibling;

          } else {

            sectPr = allSectPr[allSectPr.length - 1];

          }

        } else if (allSectPr.length > 1) {

          sectPr = allSectPr[allSectPr.length - 2];

        } else if (allSectPr.length === 1) {

          sectPr = allSectPr[0];

        }



        const nodes = Array.from(formattedBody.childNodes).filter(

          n => n.nodeType === 1 && n.nodeName !== "w:sectPr"

        );



        nodes.forEach(node => {

          const imported = templateDoc.importNode(node, true);

          if (sectPr && sectPr.parentNode === templateBody) {

            templateBody.insertBefore(imported, sectPr);

          } else {

            templateBody.appendChild(imported);

          }

        });

      }

      const serializer = new XMLSerializer();

      templateZip.file(

        "word/document.xml",

        serializer.serializeToString(templateDoc)

      );

      const finalBuffer = templateZip.generate({

        type: "nodebuffer",

        compression: "DEFLATE"

      });



      const safeName = (tabName || "Document")

        .replace(/[/\\:*?"<>|]/g, "_")

        .trim();

      req._.res.setHeader(

        "Content-Type",

        "application/vnd.openxmlformats-officedocument.wordprocessingml.document"

      );

      req._.res.setHeader(

        "Content-Disposition",

        `attachment; filename="${safeName}.docx"`

      );

      return req._.res.send(finalBuffer);

    }

    catch (error) {

      console.error("Document generation failed:", error);



      if (!req._?.res?.headersSent) {

        req.error(500, "Document generation failed: " + error.message);

      }

    }



  });


  this.on('viewTemplate', async (req) => {

    try {

      const key = req.data.key;

      if (!key) {

        return req.error(400, 'Key parameter is required.');

      }



      const { s3, bucketName } = await getObjectStoreConfig();



      const response = await s3.send(

        new GetObjectCommand({

          Bucket: bucketName,

          Key: key

        })

      );



      const chunks = [];

      for await (const chunk of response.Body) {

        chunks.push(chunk);

      }

      const buffer = Buffer.concat(chunks);



      req._.res.setHeader(

        'Content-Type',

        response.ContentType || 'application/octet-stream'

      );

      req._.res.setHeader(

        'Content-Disposition',

        `inline; filename="${key.split('/').pop()}"`

      );



      return req._.res.send(buffer);



    } catch (error) {

      console.error('[viewTemplate] Error:', error);

      return req.error(500, 'Failed to fetch template.');

    }

  });

  this.on('extractTemplateStructure', async (req) => {
    try {
      const key = req.data.key;
      if (!key) {
        return req.error(400, 'Key parameter is required.');
      }

      const { s3, bucketName } = await getObjectStoreConfig();
      const s3Response = await s3.send(
        new GetObjectCommand({
          Bucket: bucketName,
          Key: key
        })
      );

      // Convert stream to buffer
      const chunks = [];
      for await (const chunk of s3Response.Body) {
        chunks.push(chunk);
      }
      const templateBuffer = Buffer.concat(chunks);

      // Determine file type from key extension
      const extension = (key.split('.').pop() || '').toLowerCase();
      let structure = { headings: [], sections: [], rawText: '' };

      if (extension === 'docx') {
        // Extract structure from .docx using PizZip + XML parsing
        const templateZip = new PizZip(templateBuffer);
        const docXml = templateZip.file("word/document.xml");
        if (docXml) {
          const parser = new DOMParser();
          const doc = parser.parseFromString(docXml.asText());
          const paragraphs = doc.getElementsByTagName("w:p");
          const headings = [];
          let fullText = [];

          // Helper: check whether a node is inside a table (w:tbl ancestor)
          function isInsideTable(node) {
            let p = node.parentNode;
            while (p) {
              if (p.nodeName === "w:tbl") return true;
              p = p.parentNode;
            }
            return false;
          }

          for (let i = 0; i < paragraphs.length; i++) {
            const para = paragraphs[i];

            // Skip paragraphs that are inside tables — those are table cell contents
            // (field labels / placeholder values), NOT document headings/body text.
            // Including them causes the AI to duplicate template field labels as
            // separate paragraphs in the output.
            const insideTable = isInsideTable(para);

            // Get paragraph style
            const pPr = para.getElementsByTagName("w:pPr");
            let styleId = "";
            let outlineLevel = -1;

            if (pPr.length > 0) {
              const pStyle = pPr[0].getElementsByTagName("w:pStyle");
              if (pStyle.length > 0) {
                styleId = pStyle[0].getAttribute("w:val") || "";
              }
              const outlineLvl = pPr[0].getElementsByTagName("w:outlineLvl");
              if (outlineLvl.length > 0) {
                outlineLevel = parseInt(outlineLvl[0].getAttribute("w:val"), 10);
              }
            }

            // Extract text from runs
            const runs = para.getElementsByTagName("w:t");
            let paraText = "";
            for (let j = 0; j < runs.length; j++) {
              paraText += runs[j].textContent || "";
            }
            paraText = paraText.trim();

            if (!paraText) continue;

            // Only push body paragraphs (not table-cell text) into fullText,
            // so fallback heading detection doesn't turn table field-labels
            // into fake headings.
            if (!insideTable) {
              fullText.push(paraText);
            }

            // Determine heading level
            let headingLevel = 0;
            if (/^Heading(\d)$/i.test(styleId)) {
              headingLevel = parseInt(styleId.replace(/\D/g, ''), 10);
            } else if (/^TOC\d$/i.test(styleId)) {
              continue; // Skip TOC entries
            } else if (outlineLevel >= 0 && outlineLevel <= 5) {
              headingLevel = outlineLevel + 1;
            }

            // Do NOT treat text inside tables as headings, even if it happens
            // to carry a heading style (rare but possible in some templates).
            if (!insideTable && headingLevel > 0 && headingLevel <= 6) {
              headings.push({
                level: headingLevel,
                text: paraText
              });
            }
          }

          structure.headings = headings;
          structure.rawText = fullText.join('\n');

          // Fallback: if no styled headings found, detect from text patterns.
          // Use only body text (fullText already excludes table cell content).
          if (headings.length === 0 && fullText.length > 0) {
            console.log("[extractTemplateStructure] No styled headings found, using text-based detection");
            const seenHeadings = new Set();
            fullText.forEach(line => {
              const trimmed = line.trim();
              if (!trimmed || trimmed.length > 100) return;
              // Detect numbered headings like "1. Conceptual Design" or "1.1 Brief description"
              // Detect numbered headings like "1. Conceptual Design" or "1.Conceptual Design5" (with page numbers)
              const numberedMatch = trimmed.match(/^(\d+(\.\d+)*)[\.\)]?\s*(.+)$/);
              if (numberedMatch) {
                const dots = (numberedMatch[1].match(/\./g) || []).length;
                // Strip trailing page numbers (e.g., "Conceptual Design5" -> "Conceptual Design")
                let headingText = numberedMatch[3].replace(/\d+$/, "").trim();
                if (headingText.length > 2) {
                  const full = numberedMatch[1] + " " + headingText;
                  const key = full.toLowerCase();
                  if (!seenHeadings.has(key)) {
                    seenHeadings.add(key);
                    headings.push({ level: dots + 1, text: full });
                  }
                }
                return;
              }
              // Detect ALL CAPS lines as level 1 headings
              if (trimmed.length > 3 && trimmed.length < 60 && trimmed === trimmed.toUpperCase() && /[A-Z]/.test(trimmed)) {
                const key = trimmed.toLowerCase();
                if (!seenHeadings.has(key)) {
                  seenHeadings.add(key);
                  headings.push({ level: 1, text: trimmed });
                }
                return;
              }
              // Detect Title Case short lines (likely section headers)
              if (trimmed.length > 3 && trimmed.length < 60 && /^[A-Z][a-z]/.test(trimmed) && !trimmed.includes(".") && trimmed.split(" ").length <= 6) {
                // Check if it looks like a heading (no punctuation at end, title-like)
                if (!/[,;:]$/.test(trimmed) && !/^\(/.test(trimmed)) {
                  const key = trimmed.toLowerCase();
                  if (!seenHeadings.has(key)) {
                    seenHeadings.add(key);
                    headings.push({ level: 2, text: trimmed });
                  }
                }
              }
            });
            structure.headings = headings;
          }

          // Build a hierarchical section structure
          const sections = [];
          let currentSection = null;
          headings.forEach((h) => {
            if (h.level === 1) {
              currentSection = { title: h.text, level: h.level, subsections: [] };
              sections.push(currentSection);
            } else if (currentSection && h.level === 2) {
              currentSection.subsections.push({ title: h.text, level: h.level, subsections: [] });
            } else if (currentSection && currentSection.subsections.length > 0 && h.level >= 3) {
              const lastSub = currentSection.subsections[currentSection.subsections.length - 1];
              lastSub.subsections.push({ title: h.text, level: h.level });
            } else {
              sections.push({ title: h.text, level: h.level, subsections: [] });
            }
          });
          structure.sections = sections;
        }
      } else if (extension === 'xlsx' || extension === 'xls') {
        // Extract structure from Excel - get sheet names and column headers
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.load(templateBuffer);
        const sheets = [];
        workbook.eachSheet((worksheet) => {
          const sheetInfo = { name: worksheet.name, columns: [] };
          const headerRow = worksheet.getRow(1);
          headerRow.eachCell((cell) => {
            if (cell.value) {
              sheetInfo.columns.push(String(cell.value));
            }
          });
          sheets.push(sheetInfo);
        });
        structure.sections = sheets;
        structure.headings = sheets.map(s => ({ level: 1, text: s.name }));
      } else {
        // For other file types, try mammoth for text extraction
        try {
          const result = await mammoth.extractRawText({ buffer: templateBuffer });
          structure.rawText = result.value || '';
          // Try to identify headings from text (lines that are ALL CAPS or short bold-like)
          const lines = structure.rawText.split('\n').filter(l => l.trim());
          lines.forEach(line => {
            const trimmed = line.trim();
            if (trimmed.length > 2 && trimmed.length < 80 && trimmed === trimmed.toUpperCase() && /[A-Z]/.test(trimmed)) {
              structure.headings.push({ level: 1, text: trimmed });
            }
          });
        } catch (e) {
          structure.rawText = '';
        }
      }


      // Extract table structures from the docx AND figure out which
      // template section (heading) each table belongs to. Knowing the
      // owning section lets the LLM emit each table INSIDE that section,
      // so the downstream generator can map and populate the template's
      // empty tables instead of dumping loose tables at the end.
      let tables = [];
      if (extension === "docx") {
        try {
          const templateZip2 = new PizZip(templateBuffer);
          const docXml2 = templateZip2.file("word/document.xml");
          if (docXml2) {
            const parser2 = new DOMParser();
            const doc2 = parser2.parseFromString(docXml2.asText());
            const body2 = doc2.getElementsByTagName("w:body")[0];

            const extractCellText = (cell) => {
              const texts = cell.getElementsByTagName("w:t");
              let cellText = "";
              for (let t = 0; t < texts.length; t++) {
                cellText += texts[t].textContent || "";
              }
              return cellText.trim();
            };

            // Helper: heading level for a paragraph in the template body
            const getPHeadingLevel = (para) => {
              const pPr = para.getElementsByTagName("w:pPr");
              if (pPr.length === 0) return 0;
              const pStyle = pPr[0].getElementsByTagName("w:pStyle");
              if (pStyle.length > 0) {
                const styleVal = pStyle[0].getAttribute("w:val") || "";
                const m = styleVal.match(/^Heading(\d)$/i);
                if (m) return parseInt(m[1], 10);
                if (/^TOC\d/i.test(styleVal)) return 0;
              }
              const outlineLvl = pPr[0].getElementsByTagName("w:outlineLvl");
              if (outlineLvl.length > 0) {
                const lvl = parseInt(outlineLvl[0].getAttribute("w:val"), 10);
                if (lvl >= 0 && lvl <= 5) return lvl + 1;
              }
              return 0;
            };
            const getPText = (para) => {
              const runs = para.getElementsByTagName("w:t");
              let text = "";
              for (let j = 0; j < runs.length; j++) text += runs[j].textContent || "";
              return text.trim();
            };

            // Walk direct children of body in order. When we see a heading
            // paragraph, remember its text as the "current section". When
            // we see a table, tag it with the current section heading.
            let currentSectionHeading = "";
            const orderedTables = [];
            if (body2) {
              const children = Array.from(body2.childNodes).filter(n => n.nodeType === 1);
              for (const node of children) {
                if (node.nodeName === "w:p") {
                  const lvl = getPHeadingLevel(node);
                  if (lvl > 0) {
                    const txt = getPText(node);
                    if (txt) currentSectionHeading = txt;
                  }
                } else if (node.nodeName === "w:tbl") {
                  orderedTables.push({ node: node, section: currentSectionHeading });
                } else if (node.nodeName === "w:sdt") {
                  // A structured document tag may wrap a heading — try to
                  // find its inner heading text if any
                  const innerPs = node.getElementsByTagName("w:p");
                  for (let k = 0; k < innerPs.length; k++) {
                    const lvl = getPHeadingLevel(innerPs[k]);
                    if (lvl > 0) {
                      const txt = getPText(innerPs[k]);
                      if (txt) currentSectionHeading = txt;
                    }
                  }
                }
              }
            }

            // Fallback: if for some reason we didn't collect via ordered walk
            // (e.g. tables live under w:sdt wrappers), fall back to global list
            const allTables = orderedTables.length > 0
              ? orderedTables
              : Array.from(doc2.getElementsByTagName("w:tbl")).map(n => ({ node: n, section: "" }));

            for (let ti = 0; ti < allTables.length; ti++) {
              const tbl = allTables[ti].node;
              const owningSection = allTables[ti].section || "";
              const rowNodes = tbl.getElementsByTagName("w:tr");
              if (rowNodes.length === 0) continue;

              // Build full grid
              const grid = [];
              for (let ri = 0; ri < rowNodes.length; ri++) {
                const cellNodes = rowNodes[ri].getElementsByTagName("w:tc");
                const rowCells = [];
                for (let ci = 0; ci < cellNodes.length; ci++) {
                  rowCells.push(extractCellText(cellNodes[ci]));
                }
                grid.push(rowCells);
              }
              if (grid.length === 0) continue;

              const headers = grid[0].map(c => c || "");
              const dataRows = grid.slice(1);
              const nonEmptyHeader = headers.some(h => h && h.length > 0);
              if (!nonEmptyHeader) continue;

              // Column count is the max of all rows (in case of merged cells etc.)
              const colCount = Math.max(...grid.map(r => r.length));

              // ─── Classify table kind ───
              const KEY_LABEL_REGEX = /^(field\s*name|field|label|attribute|property|parameter|name|entry\s*value|value|description|entry)$/i;
              let kind = "columnar";

              if (colCount === 2) {
                const h1 = (headers[0] || "").toLowerCase();
                const h2 = (headers[1] || "").toLowerCase();
                const headerLooksLikeKV =
                  KEY_LABEL_REGEX.test(headers[0] || "") &&
                  KEY_LABEL_REGEX.test(headers[1] || "");

                let kvRowCount = 0;
                let considered = 0;
                dataRows.forEach(r => {
                  const a = (r[0] || "").trim();
                  const b = (r[1] || "").trim();
                  if (!a && !b) return;
                  considered++;
                  const bLooksEmpty = !b || /^(-|na|n\/a|tbd|pending|to\s*be\s*(added|filled|updated|provided))$/i.test(b);
                  if (a && bLooksEmpty) kvRowCount++;
                });
                const kvRatio = considered > 0 ? kvRowCount / considered : 0;

                if (headerLooksLikeKV || kvRatio >= 0.6 || (h1 && h2 && /field|name|label/.test(h1) && /value|entry|description/.test(h2))) {
                  kind = "keyValue";
                }
              }

              const keyLabels = [];
              if (kind === "keyValue") {
                dataRows.forEach(r => {
                  const label = (r[0] || "").trim();
                  if (label) keyLabels.push(label);
                });
              }

              let emptyRowCount = 0;
              dataRows.forEach(r => {
                if (r.every(c => !c || !c.trim())) emptyRowCount++;
              });

              tables.push({
                index: ti,
                headers: headers,
                kind: kind,
                columnCount: colCount,
                rowCount: dataRows.length,
                emptyRowCount: emptyRowCount,
                keyLabels: keyLabels,
                sampleRows: dataRows.slice(0, 3),
                section: owningSection
              });
            }
          }
        } catch (e) { /* ignore table extraction errors */ }
      }
      structure.tables = tables;

      // Deduplicate headings by normalized text (case-insensitive, numbering
      // stripped). The template's ToC entries and body sub-headings often
      // produce the same title under two different forms (e.g. "1.1
      // Requirement/Story Details" and "Requirement/Story Details"). Showing
      // both to the LLM makes it emit both — with cross-references like
      // "(Currently rendered above in section 1.1)". Keep only the FIRST
      // occurrence of each normalized heading.
      {
        const seen = new Set();
        const deduped = [];
        for (const h of structure.headings) {
          const key = (h.text || "")
            .replace(/^\d+(\.\d+)*[\.\)]?\s*/g, "")
            .replace(/^SECTION\s+\d+\s*[-—:.]?\s*/i, "")
            .replace(/[━─—\-:]/g, " ")
            .replace(/\s+/g, " ")
            .trim()
            .toLowerCase();
          if (!key) continue;
          if (seen.has(key)) continue;
          seen.add(key);
          deduped.push(h);
        }
        structure.headings = deduped;
      }

      // ─── Generate a prompt-friendly template description ───
      //
      // Design goal: the LLM must produce output that has EXACTLY the same
      // section headings as the template, with each table embedded inline
      // under the same section it lives under in the template. That is what
      // lets `generateDocument` map AI content into the template.
      //
      // IMPORTANT: this text is injected into the LLM's system message. Any
      // markdown heading syntax (#, ##, ###) here is dangerous — the LLM
      // tends to copy it verbatim into its own output. So we use plain,
      // bullet-style prose to describe the template rather than markdown
      // structure. The ONLY markdown headings the LLM should output are the
      // ones under "TEMPLATE HEADINGS TO USE" below.
      let templateDescription = "TEMPLATE OUTLINE INSTRUCTIONS\n";
      templateDescription += "============================\n\n";
      templateDescription +=
        "Produce your response using EXACTLY the section headings listed under TEMPLATE HEADINGS TO USE, " +
        "in that order. Use markdown heading syntax (#, ##, ...) to match the level shown. " +
        "Under each heading, write body content for that section. If a section has a table, " +
        "output the table (in markdown pipe `|` format) IMMEDIATELY under that heading and BEFORE moving to the next heading. " +
        "Instruction lines that begin with `->` or that appear inside parentheses are for you (the writer) only — do NOT include them in your output.\n\n";

      // Group tables by their owning section (normalized text)
      const tablesBySection = new Map();
      tables.forEach(t => {
        const key = (t.section || "").trim().toLowerCase();
        if (!tablesBySection.has(key)) tablesBySection.set(key, []);
        tablesBySection.get(key).push(t);
      });
      const usedTableIndices = new Set();

      const renderTableHint = (tbl) => {
        const headerRow = "| " + tbl.headers.join(" | ") + " |";
        const sepRow = "| " + tbl.headers.map(() => "---").join(" | ") + " |";
        let out = "";
        out += "    " + headerRow + "\n";
        out += "    " + sepRow + "\n";
        if (tbl.kind === "keyValue") {
          if (tbl.keyLabels.length > 0) {
            out += "    -> Two-column form table. Produce ONE data row for each field label listed below, filling ONLY the second column with the value. Do NOT print these labels as separate paragraphs.\n";
            tbl.keyLabels.forEach(lbl => {
              out += "       * " + lbl + "\n";
            });
          } else {
            out += "    -> Two-column form table — fill values for each field label from the template.\n";
          }
        } else {
          out += "    -> Columnar table — produce data rows appropriate to the content. Use the EXACT column headers shown.\n";
        }
        return out;
      };

      templateDescription += "TEMPLATE HEADINGS TO USE (in order)\n";
      templateDescription += "-----------------------------------\n";
      if (structure.headings.length > 0) {
        structure.headings.forEach((h) => {
          const prefix = '#'.repeat(h.level) + ' ';
          templateDescription += prefix + h.text + "\n";
          const key = (h.text || "").trim().toLowerCase();
          const sectionTables = tablesBySection.get(key) || [];
          sectionTables.forEach(t => {
            usedTableIndices.add(t.index);
            templateDescription += "  -> Include a table here with the following headers (do NOT print the `->` line):\n";
            templateDescription += renderTableHint(t);
          });
        });
      } else {
        templateDescription +=
          "(No headings could be detected in the template — follow a natural document structure and include the tables listed below.)\n";
      }

      // Any tables that couldn't be attached to a section — list them at the
      // end as a plain-text block (NOT a markdown heading, so the LLM does
      // not copy the label into its response).
      const orphanTables = tables.filter(t => !usedTableIndices.has(t.index));
      if (orphanTables.length > 0) {
        templateDescription += "\nADDITIONAL TABLES (place each under the most relevant heading above):\n";
        orphanTables.forEach(t => {
          templateDescription += renderTableHint(t);
          templateDescription += "\n";
        });
      }

      // Global content rules to prevent the issues seen in previous outputs.
      // These rules are prefixed with `-` so the LLM sees them as bullet
      // instructions, not as markdown headings.
      templateDescription += "\nSTRICT OUTPUT RULES\n";
      templateDescription += "-------------------\n";
      templateDescription +=
        "- Use each template heading EXACTLY ONCE. Never repeat a heading in numbered AND unnumbered form (e.g. do not write `1.1 Requirement/Story Details` AND later `Requirement/Story Details`).\n" +
        "- Never write cross-reference stubs like `(Currently rendered above in section 1.1)`, `(Repeated to ensure clarity)`, `(Displayed in section 1.2)`. Every section must contain its own real content.\n" +
        "- Do NOT output a Table of Contents. The template already contains one.\n" +
        "- Do NOT output the words `Table N`, `(keyValue)`, `(columnar)`, `Additional tables`, `Table for this section`, or any `->` line — those are instructions to you, not output.\n" +
        "- Do NOT copy the template's placeholder instructions (e.g. \"Describe the...\", \"Refer to Section X\", \"[Placeholder]\", angle-bracket hints like `<Failure Point>`).\n" +
        "- Do NOT print field labels (e.g. \"Object ID\", \"Title\", \"Version\") as standalone paragraphs; they belong inside the corresponding table rows only.\n" +
        "- Place every table IMMEDIATELY under its owning section heading, not at the end of the document.\n" +
        "- Keep the section ordering identical to the outline above.\n";

      return JSON.stringify({
        status: 200,
        structure: structure,
        templateDescription: templateDescription,
        headingCount: structure.headings.length,
        tableCount: tables.length
      });

    } catch (error) {
      console.error('[extractTemplateStructure] Error:', error);
      return JSON.stringify({
        status: 500,
        message: 'Failed to extract template structure: ' + error.message,
        structure: { headings: [], sections: [], rawText: '' },
        templateDescription: ''
      });
    }
  });


  this.on('sessionDataExcelAd', async (req) => {
    if (!req.user.is('Admin')) {
      return req.error(403, 'Access denied. Admin role required.');
    }

    try {
      const { fromDate, toDate } = req.data;

      if (!fromDate || !toDate) {
        return req.error(400, 'fromDate and toDate are required');
      }

      const normalizedFromDate = new Date(fromDate + "T00:00:00").toISOString().slice(0, 19).replace('T', ' ');
      const normalizedToDate = new Date(toDate + "T23:59:59").toISOString().slice(0, 19).replace('T', ' ');

      const result1 = await cds.run(
        SELECT.from('cockpit.user_login_details AS uld')
          .columns(
            'uld.session_id',
            'uld.login_time',
            'uld.logout_time',
            'uld.session_duration',
            'uld.tokens_consumed',
            'uld.project',
            'uld.Email_Id',
            'uld.UserName',
            'modelUsage.model_name',
            'modelUsage.tokens_used'
          )
          .leftJoin('cockpit.model_usage AS modelUsage')
          .on('uld.session_id = modelUsage.session_id')
          .where(
            `uld.login_time >= '${normalizedFromDate}' and uld.login_time <= '${normalizedToDate}'`
          )
          .orderBy({ 'uld.login_time': 'desc' })
      );

      const aggregatedSessions = await aggregateModelUsageBySession(result1);

      let userSessionData = {};

      aggregatedSessions.forEach(session => {
        if (!session.Email_Id || !session.login_time) return;

        let userKey = session.Email_Id;
        let date = new Date(session.login_time).toISOString().split("T")[0];
        let project = session.project;

        if (!userSessionData[userKey]) {
          userSessionData[userKey] = {
            USERNAME: session.UserName,
            EMAIL_ID: session.Email_Id,
            sessionHistory: {}
          };
        }

        if (!userSessionData[userKey].sessionHistory[date]) {
          userSessionData[userKey].sessionHistory[date] = {};
        }

        if (!userSessionData[userKey].sessionHistory[date][project]) {
          userSessionData[userKey].sessionHistory[date][project] = {
            totalSessions: 0,
            totalDurationMs: 0,
            totalTokensConsumed: 0,
            models: []
          };
        }

        userSessionData[userKey].sessionHistory[date][project].totalSessions += 1;

        if (session.session_duration) {
          try {
            const [hours, minutes, seconds] = session.session_duration.split(":").map(Number);
            userSessionData[userKey].sessionHistory[date][project].totalDurationMs +=
              (hours * 3600000) + (minutes * 60000) + (seconds * 1000);
          } catch (error) {
            console.error(`Invalid SESSION_DURATION format for record:`, session, error);
          }
        }

        if (session.tokens_consumed) {
          userSessionData[userKey].sessionHistory[date][project].totalTokensConsumed += Number(session.tokens_consumed);
        }
        mergeModelUsage(userSessionData[userKey].sessionHistory[date][project].models, session.models);
      });

      const formatExcelSessionEntry = (user, date, project, data) => {
        const totalMs = data.totalDurationMs;
        const hours = Math.floor(totalMs / (1000 * 60 * 60));
        const minutes = Math.floor((totalMs % (1000 * 60 * 60)) / (1000 * 60));
        const seconds = Math.floor((totalMs % (1000 * 60)) / 1000);

        return {
          USERNAME: user.USERNAME,
          EMAIL_ID: user.EMAIL_ID,
          date,
          project,
          totalSessions: data.totalSessions,
          totalDuration: `${String(hours).padStart(2, "0")}:${String(minutes).padStart(2, "0")}:${String(seconds).padStart(2, "0")}`,
          totalTokensConsumed: data.totalTokensConsumed,
          models: data.models
        };
      };

      const result = [];
      for (const user of Object.values(userSessionData)) {
        for (const [date, projects] of Object.entries(user.sessionHistory)) {
          for (const [project, data] of Object.entries(projects)) {
            result.push(formatExcelSessionEntry(user, date, project, data));
          }
        }
      }
      result.sort((a, b) => new Date(b.date) - new Date(a.date));

      const grouped = {};
      result.forEach(entry => {
        const month = new Date(entry.date).toLocaleString('default', { month: 'long', year: 'numeric' });
        if (!grouped[month]) grouped[month] = [];

        grouped[month].push({
          ...entry,
          models: entry.models
            .map(m => `${m.model_name}(${m.tokens_used})`)
            .filter(m => m && !m.startsWith('null(') && !m.endsWith('(0)'))
            .join(',')
        });
      });

      const workbook = new ExcelJS.Workbook();
      for (const [month, data] of Object.entries(grouped)) {
        const sheet = workbook.addWorksheet(month.substring(0, 31));
        sheet.columns = [
          { header: 'USERNAME', key: 'USERNAME' },
          { header: 'EMAIL_ID', key: 'EMAIL_ID' },
          { header: 'Date', key: 'date' },
          { header: 'Project', key: 'project' },
          { header: 'Total Sessions', key: 'totalSessions' },
          { header: 'Total Duration', key: 'totalDuration' },
          { header: 'Total Tokens Consumed', key: 'totalTokensConsumed' },
          { header: 'Models', key: 'models' }
        ];
        sheet.addRows(data);
        sheet.getRow(1).eachCell((cell) => {
          cell.font = { bold: true };
        });
      }

      workbook.eachSheet(sheet => {
        sheet.eachRow(row => {
          row.eachCell(cleanCellValue);
        });
      });

      const buffer = Buffer.from(await workbook.xlsx.writeBuffer());
      const base64 = buffer.toString('base64');
      const downloadDate = new Date().toISOString().split('T')[0];

      if (req._?.res && !req._.res.headersSent) {
        req._.res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        req._.res.setHeader('Content-Disposition', `attachment; filename="session_data_${downloadDate}.xlsx"`);
        req._.res.setHeader('Content-Length', buffer.length);
        return req._.res.send(buffer);
      }
      return {
        filename: `session_data_${downloadDate}.xlsx`,
        contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        data: base64
      };
    } catch (error) {
      return req.error(500, "Error generating Excel: " + error.message);
    }
  });

  this.on('promptDataExcel', async (req) => {
    try {
      const { project, fromDate, toDate } = req.data;

      if (!project || !fromDate || !toDate) {
        return req.error(400, 'Required parameters: project, fromDate, toDate');
      }

      const from = new Date(fromDate);
      const to = new Date(toDate);
      to.setHours(23, 59, 59, 999);

      const fromISO = from.toISOString();
      const toISO = to.toISOString();

      const tx = cds.transaction(req);

      const rows = await tx.run(
        `SELECT
          "user_id",
          "session_id",
          "prompt",
          "system_id",
          "sysmsg",
          "token_consumed",
          "Date_Added",
          "model_name",
          "project"
        FROM "AIcockpit"."devcockpit_Prompt_logs"
        WHERE "project" = $1
          AND "Date_Added" >= $2
          AND "Date_Added" <= $3
        ORDER BY "Date_Added" DESC`,
        [project, fromISO, toISO]
      );

      function normalizeDateToISO(value) {
        if (!value) return null;
        if (value instanceof Date && !isNaN(value)) return value.toISOString();
        if (typeof value === 'string') {
          const s = value.trim();
          if (!s) return null;
          if (s.includes('T')) {
            return /Z$|[+-]\d{2}:\d{2}$/.test(s) ? s : `${s}Z`;
          }
          const withT = s.replace(' ', 'T');
          const normalized = withT.replace(/(\.\d{3})\d+$/, '$1');
          return `${normalized}Z`;
        }
        const d = new Date(value);
        return isNaN(d) ? null : d.toISOString();
      }

      const detailRows = (rows || []).map(r => {
        const iso = normalizeDateToISO(r.Date_Added || r.date_added);
        const dateStr = iso ? new Date(iso).toISOString().split('T')[0] : '';
        return {
          user_id: r.user_id || r.USER_ID,
          date: dateStr,
          project: r.project || r.PROJECT,
          model_tokens: `${r.model_name || r.MODEL_NAME || 'Unknown'}(${Number(r.token_consumed || r.TOKEN_CONSUMED) || 0})`,
          prompt: r.prompt || r.PROMPT || '',
          system_id: r.system_id || r.SYSTEM_ID || ''
        };
      }).sort((a, b) => b.date.localeCompare(a.date));

      const workbook = new ExcelJS.Workbook();
      const sheet = workbook.addWorksheet('Details');

      sheet.columns = [
        { header: 'USER_ID', key: 'user_id', width: 24 },
        { header: 'Date', key: 'date', width: 12 },
        { header: 'Project', key: 'project', width: 22 },
        { header: 'Model Used (TokensPerModel)', key: 'model_tokens', width: 34 },
        { header: 'SystemID', key: 'system_id', width: 24 },
        { header: 'Prompt', key: 'prompt', width: 80 },
      ];

      sheet.addRows(detailRows);
      sheet.getRow(1).eachCell(cell => { cell.font = { bold: true }; });

      sheet.eachRow(row => {
        row.eachCell(cell => {
          if (cell.value) {
            const val = String(cell.value).trim();
            if (['null', 'null(null)', 'null(0)'].includes(val)) cell.value = '';
          }
        });
      });

      const buffer = await workbook.xlsx.writeBuffer();

      return Buffer.from(buffer).toString('base64');

    } catch (error) {
      console.error('Internal Server Error', error);
      return req.error(500, `Error generating Excel: ${error.message}`);
    }
  });

  this.on("uploadFile", async (req) => {
    try {
      const tx = cds.transaction(req);

      const { Category, Project, userId, fileName, mimeType, fileBase64 } = req.data.payload;

      const missingFields = [];
      if (!Category) missingFields.push("Category");
      if (!Project) missingFields.push("Project");
      if (!userId) missingFields.push("userId");
      // if (!isValidEmail(userId)) return req.reject(400, 'UserId must be a valid email address');

      // if (req.user?.id?.toLowerCase() !== userId?.toLowerCase()) {
      //   return req.reject(400, 'Mismatched email id');
      // }

      if (!fileName) missingFields.push("fileName");
      if (!mimeType) missingFields.push("mimeType");
      if (!fileBase64) missingFields.push("fileBase64");

      // if (missingFields.length > 0) {
      //   return req.error(400, "Missing mandatory fields: " + missingFields.join(", "));
      // }

      const allowedMimeTypes = [
        'application/pdf',
        'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'text/plain',
        'image/jpeg',
        'image/png'
      ];

      if (!allowedMimeTypes.includes(mimeType)) {
        return req.error(400, "File type not allowed. Only PDF, DOCX, XLSX, TXT, and image files are permitted.");
      }

      const fileExtensions = fileName.match(/\.[a-zA-Z0-9]+/g) || [];
      if (fileExtensions.length > 1) {
        return req.error(400, "Invalid file name. Multiple extensions are not allowed (e.g., .txt.docx, .txt.txt).");
      }

      const allowedExtensions = {
        'application/pdf': ['.pdf'],
        'application/vnd.openxmlformats-officedocument.wordprocessingml.document': ['.docx'],
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'],
        'text/plain': ['.txt'],
        'image/jpeg': ['.jpg', '.jpeg'],
        'image/png': ['.png']
      };
      const fileExtension = fileName.substring(fileName.lastIndexOf('.')).toLowerCase();
      const expectedExtensions = allowedExtensions[mimeType] || [];

      if (!expectedExtensions.includes(fileExtension)) {
        return req.error(400, `File extension "${fileExtension}" does not match the declared file type "${mimeType}".`);
      }

      const sanitizedFileName = fileName.replace(/[^a-zA-Z0-9._-]/g, '_');
      if (sanitizedFileName !== fileName) {
        console.log(`[SECURITY] File name sanitized from "${fileName}" to "${sanitizedFileName}"`);
      }
      const buffer = Buffer.from(fileBase64, "base64");
      const maxSize = 10 * 1024 * 1024;
      if (buffer.length > maxSize) {
        return req.error(400, "File size exceeds 10MB limit");
      }

      const contentString = buffer.toString('utf8', 0, Math.min(buffer.length, 100000));
      const maliciousPatterns = [
        /<script\b[^>]*>[\s\S]*?<\/script>/gi,
        /javascript:/gi,
        /vbscript:/gi,
        /on\w+\s*=\s*["'][^"']*["']/gi,
        /eval\s*\(/gi,
        /document\.write/gi,
        /document\.cookie/gi,
        /window\.location/gi,
        /\.exec\s*\(/gi,
        /new\s+Function\s*\(/gi,
        /fromCharCode/gi,
        /\\x[0-9a-fA-F]{2}/g,
        /\\u[0-9a-fA-F]{4}/g,
        /base64_decode/gi,
        /shell_exec/gi,
        /system\s*\(/gi,
        /passthru/gi,
        /exec\s*\(/gi,
        /popen\s*\(/gi,
        /proc_open/gi,
        /<\?php/gi,
        /<%[\s\S]*?%>/g,
        /powershell/gi,
        /cmd\.exe/gi,
        /\/bin\/sh/gi,
        /\/bin\/bash/gi,
        /wget\s+/gi,
        /curl\s+.*-o/gi,
        /nc\s+-e/gi,
        /rm\s+-rf/gi
      ];

      for (const pattern of maliciousPatterns) {
        if (pattern.test(contentString)) {
          console.warn(`[SECURITY] Malicious content detected in file "${fileName}" by user "${userId}". Pattern: ${pattern}`);
          req.warn("File contains potentially malicious content. Proceeding with upload.");
          break;
        }
      }

      const { s3, bucketName } = await getObjectStoreConfig();

      const existing = await tx.run(
        SELECT.from('devcockpit.FileDetails')
          .where({ FileName: sanitizedFileName, Category: Category, Project: Project })
      );

      if (existing.length > 0) {
        const objectStoreRefKey = existing[0].ObjectStoreRefKey;

        const putCommand = new PutObjectCommand({
          Bucket: bucketName,
          Key: objectStoreRefKey,
          Body: buffer,
          ContentType: mimeType
        });
        await s3.send(putCommand);

        return {
          status: 200,
          fileName,
          objectStoreRefKey,
          message: "File overwritten successfully"
        };
      }
      console.log("No existing file found, proceeding with new upload");
      const objectStoreRefKey = `${Project}/${Category}/${sanitizedFileName}`;
      // const objectStoreRefKey = `${cds.utils.uuid()}/${fileName}`;
      console.log("objectstore :", objectStoreRefKey)
      const maxIdRow = await tx.run(
        SELECT.from('devcockpit.FileDetails')
          .columns('coalesce(max(ID), 0) + 1 as nextid')
      );

      const nextId = maxIdRow[0]?.nextid || maxIdRow[0]?.NEXTID || 1;
      console.log('maxIdRow:', maxIdRow);
      console.log('nextId:', nextId);
      await tx.run(
        INSERT.into('devcockpit.FileDetails').entries({
          ID: nextId,
          FileName: sanitizedFileName,
          ObjectStoreRefKey: objectStoreRefKey,
          FileType: mimeType,
          CreatedBy: userId,
          Category: Category,
          Project: Project,
          Date_Added: new Date()
        })
      );


      const putCommand = new PutObjectCommand({
        Bucket: bucketName,
        Key: objectStoreRefKey,
        Body: buffer,
        ContentType: mimeType
      });
      await s3.send(putCommand);

      return {
        status: 200,
        fileName: sanitizedFileName,
        objectStoreRefKey,
        message: "File uploaded successfully"
      };

    } catch (err) {
      console.error("uploadFile error:", err);
      return req.error(500, "File upload failed");
    }
  });



  // this.on('getDeployments', async (req) => {
  //   try {
  //     const { User_Email_Id } = req.data;


  //     const apiResponse = await cds.db.tx(async tx => {
  //       return tx.run(
  //         SELECT.from('devcockpit_UserResourceMapping')
  //           .where(`LOWER(EmailID) = LOWER('${User_Email_Id}')`)
  //       );
  //     });

  //     if (apiResponse.length > 0) {
  //       const apiResult = await fetchAIModels(apiResponse[0], req);

  //       return req.reply({
  //         status: 201,
  //         message: 'Deployment details fetched successfully',
  //         result: apiResult
  //       });
  //     } else {
  //       return req.error(403, 'You do not have access. User not found.');
  //     }
  //   } catch (err) {
  //     console.error('Error in getDeployments:', err.message, err.stack);
  //     req.error({
  //       code: '500',
  //       message: `Internal Server Error: ${err.message}`,
  //       target: 'getDeployments',
  //       status: 500
  //     });
  //   }
  // });

  this.on('updateProject', async (req) => {
    const { session_id, project } = req.data;

    if (!session_id || !project) {
      return req.error(400, 'Please provide both session_id and project');
    }

    const tx = cds.transaction(req);

    try {
      const result = await tx.run(
        SELECT.from('devcockpit_user_login_details')
          .columns('project')
          .where({ session_id })
      );

      if (!result || result.length === 0) {
        return req.error(404, 'Session ID not found');
      }

      if (!result[0].project) {
        await tx.run(
          UPDATE('devcockpit_user_login_details')
            .set({ project })
            .where({ session_id })
        );

        return { status: 200, message: 'Project updated successfully' };
      } else {
        return { message: 'Project already set', currentProject: result[0].project };
      }
    } catch (error) {
      console.error("Error in updateProject:", error);
      return req.error(500, 'Internal Server Error');
    }
  });

  this.on('saveLogin', async (req) => {
    try {
      const { Email_Id, UserName, login_time } = req.data.payload;
      if (!Email_Id || !UserName || !login_time) {
        return req.reject(400, "Email_Id, UserName, login_time, and session_id are required.");
      }
      // if (!isValidEmail(Email_Id)) return req.reject(400, 'Email_Id must be a valid email address');
      // if (req.user?.id?.toLowerCase() !== Email_Id?.toLowerCase()) {
      //   return req.reject(400, 'Mismatched email id');
      // }

      const tx = cds.transaction(req);


      const sequence = await tx.run(
        SELECT.from('devcockpit_user_login_details')
          .columns('coalesce(max(ID),0) + 1 as maxid')
      );


      let maxId = parseInt(sequence[0]?.maxid || sequence[0]?.MAXID || 1);
      if (!maxId || isNaN(maxId)) maxId = 1;

      const session_id = uuidv4();

      const result = await tx.run(
        INSERT.into('devcockpit_user_login_details').entries({
          ID: maxId,
          Email_Id,
          UserName,
          login_time,
          session_id,
          tokens_consumed: 0
        })
      );


      if (result) {
        return req.reply({
          status: 201,
          message: 'Login details saved successfully',
          result: { session_id: session_id }
        });
      } else {
        return req.error(500, 'Error in saving login details');
      }
    } catch (err) {
      console.error('Error in saveLogin:', err);
      return req.error(500, 'Error in processing request');
    }
  });

  this.on('saveLogout', async (req) => {
    try {
      const { Email_Id, logout_time, session_id } = req.data.payload;
      if (!Email_Id || !logout_time || !session_id) {
        return req.reject(400, "Email_Id, logout_time, and session_id are required.");
      }
      // if (!isValidEmail(Email_Id)) return req.reject(400, 'Email_Id must be a valid email address');
      // if (req.user?.id?.toLowerCase() !== Email_Id?.toLowerCase()) {
      //   return req.reject(400, 'Mismatched email id');
      // }

      const tx = cds.transaction(req);


      const result = await tx.run(
        SELECT.from('devcockpit_user_login_details')
          .where({
            Email_Id,
            session_id,
            logout_time: null
          })
          .orderBy('login_time desc')
          .limit(1)
      );

      if (result.length === 0) {
        return req.reject(404, "No active session found for this user.");
      }

      const loginRecord = result[0];

      const sessionDurationMs = new Date(logout_time) - new Date(loginRecord.login_time);

      const hours = Math.floor(sessionDurationMs / (1000 * 60 * 60));
      const minutes = Math.floor((sessionDurationMs % (1000 * 60 * 60)) / (1000 * 60));
      const seconds = Math.floor((sessionDurationMs % (1000 * 60)) / 1000);

      const sessionDurationTime = `${String(hours).padStart(2, "0")}:${String(minutes).padStart(2, "0")}:${String(seconds).padStart(2, "0")}`;

      await tx.run(
        UPDATE('devcockpit_user_login_details')
          .set({
            logout_time,
            session_duration: sessionDurationTime
          })
          .where({ session_id })
      );

      return { message: "Logout recorded successfully", session_duration: sessionDurationTime };

    } catch (err) {
      console.error('Error in saveLogout:', err);
      return req.error(500, 'Error in processing request');
    }
  });

  this.on('getLoginDetails', async (req) => {
    try {
      const { Email_Id, project } = req.data;
      // if (!isValidEmail(Email_Id)) return req.reject(400, 'Email_Id must be a valid email address');

      const tx = cds.transaction(req);

      if (!Email_Id) {
        return req.reject(400, "Email_Id required.");
      }

      const result = await tx.run(
        SELECT.from('cockpit.user_login_details AS uld')
          .leftJoin('cockpit.model_usage AS mu')
          .on('uld.session_id = mu.session_id')
          .columns(
            'uld.session_id',
            'uld.login_time',
            'uld.logout_time',
            'uld.session_duration',
            'uld.tokens_consumed',
            'uld.project',
            'uld.Email_Id',
            'uld.UserName',
            'mu.model_name',
            'mu.tokens_used'
          )
          .where({
            'uld.Email_Id': Email_Id,
            'uld.project': project
          })
          .orderBy('uld.login_time desc')
          .limit(100)
      );


      if (!result.length) {
        return { message: "No sessions found for this user." };
      }
      const aggregatedSessions = await aggregateModelUsageBySession(result)

      let sessionDataByDateProject = {};

      aggregatedSessions.forEach(session => {
        let date = new Date(session.login_time).toISOString().split("T")[0];
        let project = session.project;
        let key = `${date}_${project}`;

        if (!sessionDataByDateProject[key]) {
          sessionDataByDateProject[key] = {
            date,
            project,
            totalSessions: 0,
            totalDurationMs: 0,
            totalTokensConsumed: 0,
            models: []
          };
        }

        sessionDataByDateProject[key].totalSessions += 1;

        if (session.session_duration) {
          const [hours, minutes, seconds] = session.session_duration.split(":").map(Number);
          sessionDataByDateProject[key].totalDurationMs += (hours * 3600000) + (minutes * 60000) + (seconds * 1000);
        }

        if (session.tokens_consumed) {
          sessionDataByDateProject[key].totalTokensConsumed += Number(session.tokens_consumed);
        }

        if (session.models) {
          session.models.forEach(({ model_name, tokens_used }) => {
            const existingModel = sessionDataByDateProject[key].models.find(m => m.model_name == model_name);
            if (existingModel) {
              existingModel.tokens_used += tokens_used;
            } else {
              sessionDataByDateProject[key].models.push({ model_name, tokens_used });
            }
          });
        }

      });


      let response = Object.entries(sessionDataByDateProject).map(([, data]) => {
        const totalMs = data.totalDurationMs;
        const hours = Math.floor(totalMs / (1000 * 60 * 60));
        const minutes = Math.floor((totalMs % (1000 * 60 * 60)) / (1000 * 60));
        const seconds = Math.floor((totalMs % (1000 * 60)) / 1000);

        const normalized = (data.models || []).map(m => ({
          model_name: (m?.model_name && String(m.model_name).trim().toLowerCase() !== 'null')
            ? String(m.model_name).trim()
            : null,
          tokens_used: Number(m?.tokens_used || 0)
        }));

        const models = normalized.length > 1
          ? normalized.filter(m => m.model_name && m.tokens_used > 0)
          : normalized;

        return {
          date: data.date,
          project: data.project,
          totalSessions: data.totalSessions,
          totalDuration: `${String(hours).padStart(2, '0')}:${String(minutes).padStart(2, '0')}:${String(seconds).padStart(2, '0')}`,
          totalTokensConsumed: data.totalTokensConsumed,
          models
        };
      });

      const finalResponse = {
        sessionHistory: response,
        EMAIL_ID: result[0].Email_Id,
        USERNAME: result[0].UserName
      };

      return {
        status: 200,
        result: finalResponse,
        count: result.length,
        message: result.length === 0 ? 'No Data found' : 'Executed'
      };

    } catch (err) {
      console.error('ERROR:', err);
      req.error({
        code: '500',
        message: `Internal Server error ${err}`,
        target: 'getLoginDetails',
        status: 500
      });
    }
  });

  this.on('getLoginDetailsOfAllUserAd', async (req) => {
    try {
      if (!req.user.is('Admin')) {
        return req.error(403, 'Access denied. Admin role required.');
      }
      const { project } = req.data;
      const tx = cds.transaction(req);


      const result1 = await tx.run(
        SELECT.from('cockpit.user_login_details AS uld')
          .leftJoin('cockpit.model_usage AS mu')
          .on('uld.session_id = mu.session_id')
          .columns(
            'uld.session_id',
            'uld.login_time',
            'uld.logout_time',
            'uld.session_duration',
            'uld.tokens_consumed',
            'uld.project',
            'uld.Email_Id',
            'uld.UserName',
            'mu.model_name',
            'mu.tokens_used'
          )
          .where({ 'uld.project': project })
          .orderBy('uld.login_time desc')
          .limit(150)
      );
      console.log('Raw result from DB:', JSON.stringify(result1.slice(0, 3), null, 2));

      const aggregatedSessions = await aggregateModelUsageBySession(result1)

      let userSessionData = {};

      aggregatedSessions.forEach(session => {
        if (!session.Email_Id || !session.login_time) return;

        let userKey = session.Email_Id;
        let date = new Date(session.login_time).toISOString().split("T")[0];
        let project = session.project

        if (!userSessionData[userKey]) {
          userSessionData[userKey] = {
            USERNAME: session.UserName,
            EMAIL_ID: session.Email_Id,
            sessionHistory: {}
          };
        }

        if (!userSessionData[userKey].sessionHistory[date]) {
          userSessionData[userKey].sessionHistory[date] = {};
        }

        if (!userSessionData[userKey].sessionHistory[date][project]) {
          userSessionData[userKey].sessionHistory[date][project] = {
            totalSessions: 0,
            totalDurationMs: 0,
            totalTokensConsumed: 0,
            models: []
          };
        }

        userSessionData[userKey].sessionHistory[date][project].totalSessions += 1;

        if (session.session_duration) {
          try {
            const [hours, minutes, seconds] = session.session_duration.split(":").map(Number);
            userSessionData[userKey].sessionHistory[date][project].totalDurationMs +=
              (hours * 3600000) + (minutes * 60000) + (seconds * 1000);
          } catch (error) {
            console.error(`Invalid SESSION_DURATION format for record:`, session);
          }
        }
        if (session.tokens_consumed) {
          userSessionData[userKey].sessionHistory[date][project].totalTokensConsumed += Number(session.tokens_consumed);
        }

        if (session.models) {
          session.models.forEach(({ model_name, tokens_used }) => {
            const existingModel = userSessionData[userKey].sessionHistory[date][project].models.find(m => m.model_name == model_name);
            if (existingModel) {
              existingModel.tokens_used += tokens_used;
            } else {
              userSessionData[userKey].sessionHistory[date][project].models.push({ model_name, tokens_used });
            }
          });
        }
      });

      let result = [];

      Object.values(userSessionData).forEach(user => {
        Object.entries(user.sessionHistory).forEach(([date, projects]) => {
          Object.entries(projects).forEach(([project, data]) => {
            let totalMs = data.totalDurationMs;
            let hours = Math.floor(totalMs / (1000 * 60 * 60));
            let minutes = Math.floor((totalMs % (1000 * 60 * 60)) / (1000 * 60));
            let seconds = Math.floor((totalMs % (1000 * 60)) / 1000);


            const normalized = (data.models || []).map(m => ({
              model_name: (m?.model_name && String(m.model_name).trim().toLowerCase() !== 'null')
                ? String(m.model_name).trim()
                : null,
              tokens_used: Number(m?.tokens_used || 0)
            }));

            const filteredModels = normalized.length > 1
              ? normalized.filter(m => m.model_name && m.tokens_used > 0)
              : normalized;

            result.push({
              USERNAME: user.USERNAME,
              EMAIL_ID: user.EMAIL_ID,
              date,
              project,
              totalSessions: data.totalSessions,
              totalDuration: `${String(hours).padStart(2, "0")}:${String(minutes).padStart(2, "0")}:${String(seconds).padStart(2, "0")}`,
              totalTokensConsumed: data.totalTokensConsumed,
              models: filteredModels
            });

          });
        });
      });

      result.sort((a, b) => new Date(b.date) - new Date(a.date));

      return {
        status: 200,
        result: result,
        count: result.length,
        message: result.length == 0 ? 'No Data found' : 'Executed'
      };

    } catch (err) {
      req.error({
        code: '500',
        message: `Internal Server error ${err}`,
        target: 'getLoginDetailsOfAllUser',
        status: 500
      });
    }
  });

  this.on('getProjectDetailsOfUser', async (req) => {
    try {
      const tx = cds.transaction(req);
      const { UserId, userName } = req.data.payload;

      if (!UserId) throw new Error('UserId is required');
      if (!isValidEmail(UserId)) return req.reject(400, 'UserId must be a valid email address');

      // if (req.user?.id?.toLowerCase() !== UserId?.toLowerCase()) {
      //   return req.reject(400, 'Mismatched email id');
      // }
      if (!userName) throw new Error('UserName is required');

      let apiResponse = await SELECT
        .from('DEVCOCKPIT_UserResourceMapping')
        .columns(['emailid', 'project_details'])
        .where`LOWER(emailid) = LOWER(${UserId})`

      let isNewRecordCreated = false;
      let isProjectAssigned = false;

      let result;

      if (!apiResponse.length || !apiResponse[0].project_details) {

        const hasViewerRole = req.user?.is('Viewer') || false;

        if (!hasViewerRole) {
          throw new Error('Role is not assigned to the user. Please contact administrator.');
        }

        const rawProjects = await tx.run(
          SELECT.from('DEVCOCKPIT_UserResourceMapping')
            .columns(['project_details'])
        );

        const uniqueProjects = [
          ...new Set(
            rawProjects
              .map(p => (p.project_details || '').trim())
              .filter(p => p)
          )
        ];

        let projectValue = null;

        if (uniqueProjects.length === 1) {
          projectValue = uniqueProjects[0];
        } else if (uniqueProjects.length === 0) {
          throw new Error('No Projects assigned to the user,Please contact your Administrator');
        } else {
          throw new Error('Multiple Projects Found,Please contact your Administrator');
        }

        if (!apiResponse.length) {

          const sequence = await tx.run(
            SELECT.one.from('DEVCOCKPIT_UserResourceMapping')
              .columns(['max(id) as maxid'])
          );

          let maxId = (parseInt(sequence.maxid) || 0) + 1;

          await tx.run(
            INSERT.into('DEVCOCKPIT_UserResourceMapping').entries({
              id: maxId,
              username: userName,
              emailid: UserId,
              project_details: projectValue,
              apiversion: '2024-02-15-preview'
            })
          );

          isNewRecordCreated = true;
        }

        if (apiResponse.length && !apiResponse[0].project_details) {

          await tx.run(
            UPDATE('DEVCOCKPIT_UserResourceMapping')
              .set({ project_details: projectValue })
              .where({ emailid: UserId })
          );

          isProjectAssigned = true;
        }

        apiResponse = await SELECT
          .from('DEVCOCKPIT_UserResourceMapping')
          .columns(['emailid', 'project_details'])
          .where({ emailid: UserId });
      }

      if (apiResponse.length) result = apiResponse[0];

      let newData;

      if (result && result.project_details) {
        const projectDetailsArray = result.project_details
          .split(',')
          .map(ele => ({
            project: ele.trim().replace(/(^')|('$)/g, '')
          }));

        newData = {
          Project_Details: projectDetailsArray,
          EmailID: result.emailid,
          Username: userName,
          UserRoles: {
            hasAdminRole: req.user?.is('Admin') || false,
            hasViewerRole: req.user?.is('Viewer') || false
          }
        };
      } else {
        newData = result;
      }

      return {
        status: 200,
        result: newData,
        ...(isNewRecordCreated && {
          message: 'New record has been created for the user'
        }),
        ...(!isNewRecordCreated && isProjectAssigned && {
          message: 'Project is being assigned to user'
        })
      };

    } catch (error) {
      return req.error({
        code: '500',
        message: error.message,
        target: 'getProjectDetailsOfUser',
        status: 500
      });
    }
  });

  this.on('getPromptDetailsofUser2_0', async (req) => {
    try {
      const { project, user_id, date_added } = req.data;
      const tx = cds.transaction(req);

      if (!project || !user_id || !date_added) {
        return req.reject(400, 'Please provide required fields: project, user_id, date_added');
      }

      const dateOnly = date_added.includes('T') ? date_added.split('T')[0] : date_added;
      const startOfDay = `${dateOnly} 00:00:00.000`;
      const endOfDay = `${dateOnly} 23:59:59.999`;

      const rows = await tx.run(
        SELECT.from('devcockpit_Prompt_logs')
          .columns(
            'user_id',
            'prompt',
            'system_id',
            'token_consumed',
            'Date_Added',
            'model_name',
            'project'
          )
          .where({
            project,
            user_id,
            Date_Added: { between: startOfDay, and: endOfDay }
          })
          .orderBy('Date_Added desc')
          .limit(100)
      ) || [];


      const promptRows = rows.map(r => ({
        user_id: r.user_id,
        prompt: r.prompt,
        system_id: r.system_id,
        token_consumed: Number(r.token_consumed) || 0,
        date_added: r.date_added,
        model_name: r.model_name || '(unknown)',
        project: r.project
      }));

      const totalsByModel = promptRows.reduce((acc, r) => {
        const m = r.model_name || '(unknown)';
        acc[m] = (acc[m] || 0) + r.token_consumed;
        return acc;
      }, {});

      const enrichedRows = promptRows.map(r => ({
        ...r,
        modelToken: { [r.model_name]: totalsByModel[r.model_name] || 0 }
      }));

      return {
        status: 200,
        result: {
          rows: enrichedRows
        },
        count: enrichedRows.length,
        message: enrichedRows.length
          ? 'Executed (latest 100 rows for the day)'
          : 'No Data found'
      };

    } catch (err) {
      req.error({
        code: '500',
        message: `Internal Server Error: ${err.message || err} `,
        target: 'getPromptDetailsofUser2_0',
        status: 500
      });
    }
  });

  this.on('logTokenUsage', async (req) => {
    const { session_id, model_id, model_name, tokensGenerated, user_id, Prompt, sysmsg, date_added, system_id, project } = req.data.payload;

    if (!session_id || !model_id || !tokensGenerated || !model_name || !user_id || !sysmsg || !date_added || !system_id) {
      return req.error(400, 'Please provide required data');
    }

    const tx = cds.transaction(req);


    const existingUsage = await tx.run(
      SELECT.from('devcockpit_model_usage')
        .where({ session_id, model_id })
    );

    if (existingUsage.length > 0) {

      await tx.run(
        UPDATE('devcockpit_model_usage')
          .set({
            tokens_used: { '+=': tokensGenerated }
          })
          .where({ session_id, model_id }))

    } else {

      const sequence = await tx.run(
        SELECT.from('devcockpit_model_usage')
          .columns('coalesce(max(ID),0) + 1 as maxid')
      );


      let maxId = parseInt(sequence[0]?.maxid || sequence[0]?.MAXID || 1);
      if (!maxId || isNaN(maxId)) maxId = 1;


      await tx.run(
        INSERT.into('devcockpit_model_usage').entries({
          ID: maxId,
          session_id,
          model_id,
          model_name,
          tokens_used: tokensGenerated
        })
      );
    }



    const sequencePrompt = await tx.run(
      SELECT.from('devcockpit_Prompt_logs')
        .columns('coalesce(max(ID),0) + 1 as maxid')
    );


    let maxId1 = parseInt(sequencePrompt[0]?.maxid || sequencePrompt[0]?.MAXID || 1);
    if (!maxId1 || isNaN(maxId1)) maxId1 = 1;


    await tx.run(
      INSERT.into('devcockpit_Prompt_logs').entries({
        ID: maxId1,
        Date_Added: date_added,
        session_id,
        project,
        prompt: Prompt,
        system_id,
        sysmsg,
        token_consumed: tokensGenerated,
        model_name,
        user_id
      })
    );


    await tx.run(
      UPDATE('devcockpit_user_login_details')
        .set({
          tokens_consumed: { '+=': tokensGenerated }
        })
        .where({ session_id })
    );


    return { status: 200, message: 'Token usage updated successfully' };
  });



  this.on('getFiles', async (req) => {
    try {
      const { Category, Project } = req.data;

      if (!Category || !Project) {
        return req.error({
          code: '400',
          message: 'Please provide Category and Project details',
          status: 400
        });
      }

      const tx = cds.transaction(req);

      const entity = 'devcockpit.FileDetails';
      const dbResult = await tx.run(
        SELECT.from(entity)
          .columns('FileName', 'ObjectStoreRefKey', 'FileType', 'Category', 'Project', 'CreatedBy', 'Date_Added')
          .where({ Category, Project })
      );

      if (!dbResult || dbResult.length === 0) {
        return {
          status: 201,
          message: 'No files added for the selected category',
          data: []
        };
      }

      const dbKeysSet = new Set(
        dbResult.map(item => item.ObjectStoreRefKey || item.objectstorerefkey)
      );

      const { s3, bucketName } = await getObjectStoreConfig();

      // Verify existence of each DB key directly in the object store (no prefix dependency)
      const matchedFiles = [];
      for (const row of dbResult) {
        const key = row.ObjectStoreRefKey || row.objectstorerefkey;
        if (!key) continue;
        try {
          const head = await s3.send(new HeadObjectCommand({ Bucket: bucketName, Key: key }));
          matchedFiles.push({
            Key: key,
            LastModified: head?.LastModified,
            Size: head?.ContentLength,
            ETag: head?.ETag,
            fileName: row.FileName || row.filename,
            category: row.Category || row.category,
            project: row.Project || row.project
          });
        } catch (err) {
          // Skip keys that do not exist; surface unexpected errors
          const status = err?.$metadata?.httpStatusCode;
          if (status === 404 || err?.name === 'NotFound' || err?.Code === 'NotFound') {
            continue;
          }
          throw err;
        }
      }

      if (matchedFiles.length === 0) {
        return {
          status: 201,
          message: 'No files found in object store for the selected category',
          data: []
        };
      }

      return {
        status: 200,
        message: 'Files retrieved',
        data: matchedFiles
      };

    } catch (e) {
      console.error('Error in getFiles:', e);
      return req.error({
        code: '500',
        message: 'Internal Server Error: ' + e.message,
        status: 500
      });
    }
  });


  this.on('createFeedback', async (req) => {
    try {
      const { IssueTitle, IssueDetail, Priority, IssueType, UserId, DateTime } = req.data.payload;
      // if (req.user?.id?.toLowerCase() !== UserId?.toLowerCase()) {
      //   return req.reject(400, 'Mismatched email id');
      // }
      if (!isValidEmail(UserId)) return req.reject(400, 'UserId must be a valid email address');

      // if (req.user?.id?.toLowerCase() !== UserId?.toLowerCase()) {
      //   return req.reject(400, 'Mismatched email id');
      // }

      const tx = cds.transaction(req);


      const queryResult = await tx.run(
        SELECT.from('devcockpit_feedback')
          .columns('coalesce(max(Issue_ID),0) + 1 as maxId')
      );


      let maxId = parseInt(queryResult[0]?.maxId || queryResult[0]?.maxid || 1);
      if (!maxId || isNaN(maxId)) maxId = 1;


      const result = await tx.run(
        INSERT.into('devcockpit_feedback').entries({
          Issue_ID: maxId,
          IssueTitle,
          IssueDetail,
          Priority,
          IssueType,
          CreatedBy: UserId,
          CreatedAt: DateTime,
          IssueStatus: 'Open'
        })
      );


      if (result) {
        return req.reply({
          status: 201,
          message: 'Issue created successfully',
          result: result
        });
      } else {
        return req.error(500, 'Error in creating issue');
      }
    } catch (err) {
      console.error('Error in createFeedback:', err);
      return req.error(500, 'Error in processing request');
    }
  });

  this.on('getAllFeedback', async (req) => {
    try {
      if (!req.user.is('SuperAdmin')) {
        return req.error(403, 'Access denied. SuperAdmin role required.');
      }

      const result = await cds.run(
        `SELECT "Issue_ID","IssueTitle","IssueDetail","IssueType","Priority","IssueStatus","CreatedBy","CreatedAt"
       FROM "devcockpit_feedback"
       ORDER BY "CreatedAt" DESC`
      );

      return JSON.stringify({
        status: 200,
        result: result,
        count: result.length,
        message: result.length === 0 ? 'No feedback found' : 'Executed'
      });
    } catch (err) {
      console.error('Error in getAllFeedback:', err);
      return req.reject(500, `Error processing request: ${err.message}`);
    }
  });

  this.on('promptTemplatesExcel', async (req) => {
    try {
      const jwt = req.headers?.authorization?.split(' ')[1];
      if (!jwt) {
        console.warn('promptTemplatesExcel: No JWT found in Authorization header');
      }
      const destination = await getDestination({ destinationName: 'AI_CORE_CGAI_COCKPIT', jwt: jwt });

      if (!destination || !destination.url) {
        throw new Error('AI_Core destination not found or missing URL');
      }

      const config = destination.originalProperties || {};
      const tokenServiceURL = config.tokenServiceURL;
      const clientId = config.clientId;
      const clientSecret = config.clientSecret;

      if (!tokenServiceURL || !clientId || !clientSecret) {
        throw new Error('Destination missing tokenServiceURL/clientId/clientSecret');
      }

      const token = await getAICoreToken({ tokenServiceURL, clientId, clientSecret });

      const headers = {
        'AI-Resource-Group': 'default',
        'Authorization': `Bearer ${token}`,
        'Content-Type': 'application/json',
      };

      const promptTemplatesUrl = `${destination.url}/lm/prompttemplates`;
      const response = await axios.get(promptTemplatesUrl, { headers });

      const promptTemplatesList = response.data?.resources || response.data || [];

      if (!Array.isArray(promptTemplatesList)) {
        throw new Error('Unexpected response format from /lm/prompttemplates');
      }

      const promptTemplates = [];

      const baseUrl = destination.url.endsWith('/v2')
        ? destination.url
        : (destination.url.includes('/v2') ? destination.url : `${destination.url}/v2`);

      const batchSize = 10;
      const delayBetweenBatches = 100;

      for (let i = 0; i < promptTemplatesList.length; i += batchSize) {
        const batch = promptTemplatesList.slice(i, i + batchSize);

        const batchPromises = batch.map(async (pt) => {
          try {
            const templateId = pt.id;
            if (!templateId) {
              return pt;
            }

            const detailUrl = `${baseUrl}/lm/promptTemplates/${templateId}`;

            const detailResponse = await axios.get(detailUrl, { headers, timeout: 10000 });
            const detailedTemplate = detailResponse.data || pt;

            return detailedTemplate;
          } catch (detailErr) {
            return pt;
          }
        });

        const batchResults = await Promise.all(batchPromises);
        promptTemplates.push(...batchResults);

        if (i + batchSize < promptTemplatesList.length) {
          await delay(delayBetweenBatches);
        }
      }

      const workbook = new ExcelJS.Workbook();
      const sheet = workbook.addWorksheet('PromptTemplates');

      const columns = [
        { header: 'ID', key: 'ID' },
        { header: 'uuid', key: 'uuid' },
        { header: 'Prompt_Template', key: 'Prompt_Template' },
        { header: 'Prompt_Details', key: 'Prompt_Details' },
        { header: 'Date_Added', key: 'Date_Added' },
        { header: 'Category', key: 'Category' },
        { header: 'MsgType', key: 'MsgType' },
        { header: 'UpdatedBy', key: 'UpdatedBy' },
        { header: 'UpdatedAt', key: 'UpdatedAt' },
        { header: 'CreatedBy', key: 'CreatedBy' },
        { header: 'Project_Id', key: 'Project_Id' },
        { header: 'PromptId', key: 'PromptId' },
      ];

      sheet.columns = columns;

      const todayDate = new Date().toISOString();

      promptTemplates.forEach(pt => {
        const templateMessages = pt.spec?.template || pt.template || [];

        const defaults = pt.spec?.defaults || pt.defaults || {};

        const createdBy = defaults.UserId || defaults.userId || '';
        const projectId = defaults.ProjectId || defaults.projectId || '';

        const category = pt.scenario || pt.Scenario || '';

        const promptId = pt.name || pt.Name || '';

        if (Array.isArray(templateMessages) && templateMessages.length > 0) {
          templateMessages.forEach(msg => {
            const role = msg.role || '';
            const content = msg.content || '';

            let msgType = '';
            if (role.toLowerCase() === 'user') {
              msgType = 'prompt';
            } else if (role.toLowerCase() === 'system') {
              msgType = 'sysMsg';
            } else {
              msgType = role;
            }

            const row = {
              ID: uuidv4(),
              Prompt_Details: content,
              Date_Added: todayDate,
              Category: category,
              MsgType: msgType,
              UpdatedBy: '',
              UpdatedAt: '',
              CreatedBy: createdBy,
              Project_Id: projectId,
              PromptId: promptId,
            };

            sheet.addRow(row);
          });
        } else {

          const row = {

            ID: uuidv4(),
            Prompt_Details: '',
            Date_Added: todayDate,
            Category: category,
            MsgType: '',
            UpdatedBy: '',
            UpdatedAt: '',
            CreatedBy: createdBy,
            Project_Id: projectId,
            PromptId: promptId,
          };

          sheet.addRow(row);
        }
      });

      sheet.getRow(1).eachCell((cell) => {
        cell.font = { bold: true };
      });

      sheet.eachRow((row, rowNumber) => {
        if (rowNumber === 1) return;
        row.eachCell((cell) => {
          if (cell.value) {
            const val = String(cell.value).trim();
            if (['null', 'undefined', 'null(null)', 'null(0)'].includes(val.toLowerCase())) {
              cell.value = '';
            }
          }
        });
      });

      sheet.views = [{ state: 'frozen', ySplit: 1 }];

      const buffer = await workbook.xlsx.writeBuffer();

      req._.res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
      req._.res.setHeader('Content-Disposition', 'attachment; filename="PromptTemplates.xlsx"');

      return req._.res.send(Buffer.from(buffer));

    } catch (error) {
      console.error('Error in promptTemplatesExcel:', error);
      return req.error(500, `Error generating Excel: ${error.message}`);
    }
  });

  this.on('migrateFilePaths', async (req) => {
    try {
      const tx = cds.transaction(req);
      const { s3, bucketName } = await getObjectStoreConfig();

      // Fetch all file records from DB
      const allFiles = await tx.run(
        `SELECT "ID", "FileName", "ObjectStoreRefKey", "Category", "Project"
         FROM "AIcockpit"."devcockpit_FileDetails"`
      );

      if (!allFiles || allFiles.length === 0) {
        return JSON.stringify({ status: 200, message: 'No files found in database', migrated: 0, skipped: 0, errors: [] });
      }

      // UUID-based key pattern: xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx/<filename>
      const uuidPattern = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}\//i;

      const toMigrate = allFiles.filter(f => {
        const key = f.ObjectStoreRefKey || f.objectstorerefkey;
        return key && uuidPattern.test(key);
      });

      if (toMigrate.length === 0) {
        return JSON.stringify({ status: 200, message: 'All files already use the new path format. Nothing to migrate.', migrated: 0, skipped: 0, errors: [] });
      }

      console.log(`[migrateFilePaths] Found ${toMigrate.length} file(s) to migrate`);

      const results = { migrated: 0, skipped: 0, errors: [] };

      for (const file of toMigrate) {
        const oldKey = file.ObjectStoreRefKey || file.objectstorerefkey;
        const fileName = file.FileName || file.filename;
        const category = file.Category || file.category;
        const project = file.Project || file.project;

        if (!project || !category || !fileName) {
          console.warn(`[migrateFilePaths] Skipping record with missing project/category/fileName. oldKey: ${oldKey}`);
          results.skipped++;
          results.errors.push({ oldKey, error: 'Missing project, category, or fileName in DB record' });
          continue;
        }

        const newKey = `${project}/${category}/${fileName}`;

        // Skip if old and new keys are already the same
        if (oldKey === newKey) {
          results.skipped++;
          continue;
        }

        try {
          console.log(`[migrateFilePaths] Copying S3 object: "${oldKey}" -> "${newKey}"`);

          // Copy object to new path in S3
          // CopySource must be "bucket/key" — only encode special chars in the key, NOT the slash
          await s3.send(new CopyObjectCommand({
            Bucket: bucketName,
            CopySource: `${bucketName}/${oldKey.split('/').map(encodeURIComponent).join('/')}`,
            Key: newKey
          }));

          // Update ObjectStoreRefKey in DB
          await tx.run(
            `UPDATE "AIcockpit"."devcockpit_FileDetails" SET "ObjectStoreRefKey" = $1 WHERE "ObjectStoreRefKey" = $2`,
            [newKey, oldKey]
          );

          // Delete old S3 object
          await s3.send(new DeleteObjectsCommand({
            Bucket: bucketName,
            Delete: { Objects: [{ Key: oldKey }] }
          }));

          console.log(`[migrateFilePaths] Successfully migrated: "${oldKey}" -> "${newKey}"`);
          results.migrated++;
        } catch (err) {
          console.error(`[migrateFilePaths] Failed to migrate "${oldKey}":`, err.message);
          results.errors.push({ oldKey, newKey, error: err.message });
          results.skipped++;
        }
      }

      const message = `Migration complete. Migrated: ${results.migrated}, Skipped/Errors: ${results.skipped} out of ${toMigrate.length} file(s).`;
      console.log(`[migrateFilePaths] ${message}`);

      return JSON.stringify({
        status: 200,
        message,
        migrated: results.migrated,
        skipped: results.skipped,
        errors: results.errors
      });

    } catch (err) {
      console.error('[migrateFilePaths] Unexpected error:', err);
      return req.error(500, 'Migration failed: ' + err.message);
    }



  });

  const URM_ENTITY = 'DEVCOCKPIT_UserResourceMapping';

  this.on('getAllUsers', async (req) => {
    if (!req.user.is('Admin')) {
      return req.error(403, 'Access denied. Admin role required.');
    }
    try {
      const tx = cds.transaction(req);
      const rows = await tx.run(
        SELECT.from(URM_ENTITY)
          .columns('ID', 'Username', 'EmailID', 'Project_Details')
          .orderBy('ID asc')
      );
      return rows || [];
    } catch (error) {
      console.error('Error in getAllUsers:', error);
      return req.error(500, 'Failed to fetch users');
    }
  });

  this.on('addUser', async (req) => {
    if (!req.user.is('Admin')) {
      return req.error(403, 'Access denied. Admin role required.');
    }
    const tx = cds.transaction(req);
    try {
      const { username, emailId, projectDetails } = req.data.payload;

      if (!username || typeof username !== 'string')
        return req.error(400, 'Username is required');
      const trimmedUsername = username.trim();
      if (!trimmedUsername)
        return req.error(400, 'Username cannot be empty');

      if (!emailId || typeof emailId !== 'string')
        return req.error(400, 'Email ID is required');
      const trimmedEmail = emailId.trim().toLowerCase();
      const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
      if (!emailRegex.test(trimmedEmail))
        return req.error(400, 'Invalid email format');

      const trimmedProjectDetails =
        typeof projectDetails === 'string' && projectDetails.trim()
          ? projectDetails.trim()
          : null;

      // Check if user already exists
      const existing = await tx.run(
        SELECT.from(URM_ENTITY)
          .columns('ID', 'Username', 'EmailID', 'Project_Details')
          .where({ EmailID: trimmedEmail })
      );
      const existingUser = existing[0] || null;

      if (existingUser) {
        if (!trimmedProjectDetails) return existingUser;

        const projects = (existingUser.Project_Details || '')
          .split(',').map(p => p.trim()).filter(Boolean);

        if (!projects.includes(trimmedProjectDetails)) {
          projects.push(trimmedProjectDetails);
          await tx.run(
            UPDATE(URM_ENTITY)
              .set({ Project_Details: projects.join(',') })
              .where({ ID: existingUser.ID })
          );

          const updated = await tx.run(
            SELECT.from(URM_ENTITY)
              .columns('ID', 'Username', 'EmailID', 'Project_Details')
              .where({ ID: existingUser.ID })
          );
          return updated[0];
        } else {
          return req.info(200, `Project '${trimmedProjectDetails}' is already assigned to user '${existingUser.Username}'. No update was made.`);
        }
      }

      // Get next ID using your working pattern
      const queryResult = await tx.run(
        SELECT.from(URM_ENTITY)
          .columns('coalesce(max(ID),0) + 1 as maxId')
      );
      let newId = parseInt(queryResult[0]?.maxId || queryResult[0]?.maxid || 1);
      if (!newId || isNaN(newId)) newId = 1;

      // Insert new user
      await tx.run(
        INSERT.into(URM_ENTITY).entries({
          ID: newId,
          Username: trimmedUsername,
          EmailID: trimmedEmail,
          Project_Details: trimmedProjectDetails
        })
      );

      const created = await tx.run(
        SELECT.from(URM_ENTITY)
          .columns('ID', 'Username', 'EmailID', 'Project_Details')
          .where({ ID: newId })
      );
      return created[0];

    } catch (err) {
      console.error('Error in addUser:', err);
      return req.error(500, err.message || 'Internal server error');
    }
  });

  this.on('deleteUser', async (req) => {
    if (!req.user.is('Admin')) {
      return req.error(403, 'Access denied. Admin role required.');
    }
    const tx = cds.transaction(req);
    try {
      const { emailId } = req.data;

      if (!emailId || typeof emailId !== 'string')
        return req.error(400, 'Email ID is required');
      const trimmedEmail = emailId.trim().toLowerCase();
      if (!trimmedEmail)
        return req.error(400, 'Email ID cannot be empty');
      const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
      if (!emailRegex.test(trimmedEmail))
        return req.error(400, 'Invalid email format');

      // Check if user exists
      const rows = await tx.run(
        SELECT.from(URM_ENTITY)
          .columns('ID')
          .where({ EmailID: trimmedEmail })
      );
      if (!rows || rows.length === 0)
        return req.error(404, `User with email '${trimmedEmail}' not found`);

      // Delete user
      await tx.run(
        DELETE.from(URM_ENTITY).where({ EmailID: trimmedEmail })
      );

      console.log(`User deleted successfully: ${trimmedEmail}`);
      return true;

    } catch (error) {
      console.error('Error in deleteUser:', error);
      return req.error(500, error.message || 'Internal server error');
    }
  });

  this.on('updateUser', async (req) => {
    if (!req.user.is('Admin')) {
      return req.error(403, 'Access denied. Admin role required.');
    }
    const tx = cds.transaction(req);
    try {
      const { id, emailId, newEmailId, username, projectDetails } = req.data.payload;

      const hasId = Number.isInteger(id) && id > 0;
      const hasEmail = typeof emailId === 'string' && emailId.trim();
      if (!hasId && !hasEmail)
        return req.error(400, 'Either ID or Email ID is required to identify the user');

      // Fetch existing user - use lowercase column names
      let existingRows;
      if (hasId) {
        existingRows = await tx.run(
          SELECT.from(URM_ENTITY)
            .columns('id', 'username', 'emailid', 'project_details')  // lowercase
            .where({ id: id })  // lowercase
        );
      } else {
        const trimmedEmail = emailId.trim().toLowerCase();
        existingRows = await tx.run(
          SELECT.from(URM_ENTITY)
            .columns('id', 'username', 'emailid', 'project_details')  // lowercase
            .where({ emailid: trimmedEmail })  // lowercase
        );
      }

      const existingUser = existingRows?.[0];
      if (!existingUser) return req.error(404, 'User not found');
      const userId = existingUser.id;  // lowercase

      // Build update object dynamically - use lowercase keys
      const updateFields = {};

      if (username !== undefined) {
        if (typeof username !== 'string') return req.error(400, 'Username must be a string');
        const t = username.trim();
        if (!t) return req.error(400, 'Username cannot be empty');
        if (t.length > 100) return req.error(400, 'Username cannot exceed 100 characters');
        updateFields.username = t;  // lowercase
      }

      if (newEmailId !== undefined) {
        if (typeof newEmailId !== 'string') return req.error(400, 'New Email ID must be a string');
        const t = newEmailId.trim().toLowerCase();
        if (!t) return req.error(400, 'New Email ID cannot be empty');
        if (t.length > 320) return req.error(400, 'New Email ID cannot exceed 320 characters');
        const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
        if (!emailRegex.test(t)) return req.error(400, 'Invalid new email format');

        const clash = await tx.run(
          SELECT.from(URM_ENTITY)
            .columns('id')  // lowercase
            .where({ emailid: t })  // lowercase
        );
        if (clash?.length > 0 && clash[0].id !== userId)  // lowercase
          return req.error(409, `Email '${t}' is already in use`);

        updateFields.emailid = t;  // lowercase
      }

      if (projectDetails !== undefined) {
        if (projectDetails !== null && typeof projectDetails !== 'string')
          return req.error(400, 'Project details must be a string');
        const t = typeof projectDetails === 'string' ? projectDetails.trim() : null;
        if (t && t.length > 320) return req.error(400, 'Project details cannot exceed 320 characters');
        updateFields.project_details = t;  // lowercase with underscore
      }

      if (Object.keys(updateFields).length === 0)
        return req.error(400, 'No fields provided for update');

      // Perform update
      await tx.run(
        UPDATE(URM_ENTITY)
          .set(updateFields)
          .where({ id: userId })  // lowercase
      );

      const updatedRows = await tx.run(
        SELECT.from(URM_ENTITY)
          .columns('id', 'username', 'emailid', 'project_details')  // lowercase
          .where({ id: userId })  // lowercase
      );
      return updatedRows[0];

    } catch (error) {
      console.error('Error in updateUser:', error);
      return req.error(500, error.message || 'Internal server error');
    }
  });

  ;

});