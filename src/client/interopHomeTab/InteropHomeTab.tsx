import * as React from "react";
import { List } from "@fluentui/react-northstar";
import { ImageFit, Pivot, PivotItem, Image} from 'office-ui-fabric-react';
import { useState, useEffect, useCallback, useRef } from "react";
import { useTeams } from "msteams-react-base-component";
import { app, authentication } from "@microsoft/teams-js";
import jwtDecode from "jwt-decode";
import {Providers, ProviderState, LoginType} from '@microsoft/mgt-element';
import * as MicrosoftTeams from "@microsoft/teams-js";
import {TeamsMsal2Provider, HttpMethod} from '@microsoft/mgt-teams-msal2-provider';
import { Msal2Provider } from '@microsoft/mgt-msal2-provider';
import App from './App';
import ReusableApp from "./ReusableApp";
import * as msal from "@azure/msal-browser";
import { Agenda, Login, FileList, Get, MgtTemplateProps, PeoplePicker, Person, ViewType} from '@microsoft/mgt-react';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import SPOReusable from "./SPOReusable";
import MSGReusable from "./MSGReusable";
import ShowAdaptiveCard from "./ShowAdaptiveCard";
/**
 * Implementation of the Interop Home content page
 */
let currentAccount: msal.AccountInfo = null;

const msalConfig = {  
    auth: {  
      clientId: 'c613e0d1-161d-4ea0-9db4-0f11eeabc2fd',
      authority: "https://login.microsoftonline.com/m365x229910.onmicrosoft.com",
      redirectUri: 'https://sridev.ngrok.io/interopHomeTab/'
    },
    cache: {
        cacheLocation: "sessionStorage", // This configures where your cache will be stored
        storeAuthStateInCookie: true, // Set this to "true" if you are having issues on IE11 or Edge
    }  
  };
  
const msalInstance = new msal.PublicClientApplication(msalConfig);

const tokenrequest: msal.SilentRequest = {
    scopes: ["https://m365x229910.sharepoint.com/AllSites.Read", "https://m365x229910.sharepoint.com/AllSites.Manage"],
    account: currentAccount,
    };

const loginRequest = {
    scopes: ["https://m365x229910.sharepoint.com/AllSites.Read", "https://m365x229910.sharepoint.com/AllSites.Manage"]
  };

export const InteropHomeTab = () => {

    const [{ inTeams, theme, context }] = useTeams();
    const [entityId, setEntityId] = useState<string | undefined>();
    const [name, setName] = useState<string>();
    const [error, setError] = useState<string>();
    const [ssoToken, setSsoToken] = useState<string>();
    const [msGraphOboToken, setMsGraphOboToken] = useState<string>();
    const [spoOboToken, setSPOOboToken] = useState<string>();
    const [searchResults, setSearchResults] = useState<any[]>();

    useEffect(() => {
        if (inTeams === true) {
            authentication.getAuthToken({
                resources: [process.env.TAB_APP_URI as string],
                silent: false
            } as authentication.AuthTokenRequestParameters).then(token => {
                const decoded: { [key: string]: any; } = jwtDecode(token) as { [key: string]: any; };
                setName(decoded!.name);
                setSsoToken(token);
                app.notifySuccess();
                TeamsMsal2Provider.microsoftTeamsLib = MicrosoftTeams;
                Providers.globalProvider = new TeamsMsal2Provider({
                    clientId: `c613e0d1-161d-4ea0-9db4-0f11eeabc2fd`,
                    authPopupUrl: '/auth.html',//'/interopHomeTab/mgtAuth.html',//
                    scopes: ['User.Read'], //,'Mail.ReadBasic'], //,'Sites.ReadWrite.All', 'Files.ReadWrite.All', 'Mail.Read'],
                    ssoUrl: 'http://localhost:5000/api/token',
                    httpMethod: HttpMethod.POST
                  });
            }).catch(message => {
                setError(message);
                app.notifyFailure({
                    reason: app.FailedReason.AuthFailed,
                    message
                });
                console.info("Error getting auth token: " + message);
            });
        } else {
            setEntityId("Not in Microsoft Teams");
            console.info("Not in Microsoft Teams");
            Providers.globalProvider = new Msal2Provider({
              clientId: 'c613e0d1-161d-4ea0-9db4-0f11eeabc2fd', //'ba686da8-8cb8-4e41-9765-056a10dee34c',
              scopes: ['User.Read'], //['calendars.read', 'user.read', 'openid', 'profile', 'people.read', 'user.readbasic.all', 'files.read', 'files.read.all'],
              redirectUri: 'https://sridev.ngrok.io/interopHomeTab/',
              loginType: LoginType.Popup
              });
            
            if(ssoToken){}
            else{

              if( msalInstance.getAllAccounts().length > 0 ) {
                  tokenrequest.account = msalInstance.getAllAccounts()[0];
              }
              msalInstance.acquireTokenSilent(tokenrequest).then((val) => {  
                let headers = new Headers();  
                let bearer = "Bearer " + val.accessToken;  
                console.info("BEARER TOKEN: "+ val.accessToken);
                console.info("ID TOKEN: "+ val.idToken);
                setSsoToken(val.idToken);
                }).catch((errorinternal) => {  
                    console.info("Internal error: "+ errorinternal); 
                    /*
                    msalInstance.loginRedirect(loginRequest).catch(e => {
                      console.log("Login error: "+ e);
                  });*/
                  msalInstance.loginPopup(loginRequest).catch(e => {
                    console.info(e);
                  });
                });
            }
            
            /*
            msalInstance.loginPopup(loginRequest).then(resp =>{
              if (resp !== null) {
                tokenrequest.account = resp.account;
                msalInstance.acquireTokenSilent(tokenrequest).then((val) => {  
                  let headers = new Headers();  
                  let bearer = "Bearer " + val.accessToken;  
                  console.info("BEARER TOKEN: "+ val.accessToken);
                  console.info("ID TOKEN: "+ val.idToken);
                  }).catch((errorinternal) => {  
                      console.log(errorinternal);  
                    });
              } 
              else{
                console.info("No account");
              }
            }).catch(function (error) {
                console.log("MSAL Login Failure: "+ error);
            }); */
        }
    }, [inTeams]);

    useEffect(() => {
        if (context) {
            setEntityId(context.page.id);
        }
    }, [context]);

    const SiteResult = (props: MgtTemplateProps) => {
      const site = props.dataContext as MicrosoftGraph.Site;
  
      return (
          <div>
              <h1>{site.name}</h1>
              {site.webUrl}
        </div>
        );
      };
    /*
    const getSPOAccessTokenOBO = async () => {
      const response = await fetch(`https://azfun.ngrok.io/api/TeamsOBOHelper?ssoToken=${ssoToken}&tokenFor=spo`);
      const responsePayload = await response.json();
      if (response.ok) {
        setSPOOboToken(responsePayload.access_token);
      } else {
        if (responsePayload!.error === "consent_required") {
          setError("consent_required");
        } else {
          setError("unknown SSO error");
        }
      }
    };

    useEffect(() => {
      // if the SSO token is defined...
      if (ssoToken && ssoToken.length > 0) {
        getSPOAccessTokenOBO();
      }
    }, [ssoToken]);      

  useEffect(() => {
    getSPOSearchResutls();
    }, [spoOboToken]);

    const getSPOSearchResutls = async () => {
      if (!spoOboToken) { return; }
    
      const endpoint = `https://m365x229910.sharepoint.com/_api/search/query?querytext=%27*%27&selectproperties=%27Author,Path,Title,Url%27&rowlimit=10`;
      const requestObject = {
        method: 'GET',
        headers: {
          "authorization": "bearer " + spoOboToken,
          "accept": "application/json; odata=nometadata"
        }
      };
    
      const response = await fetch(endpoint, requestObject);
      const responsePayload = await response.json();
    
      console.log(responsePayload.value);
      if (response.ok) {
          const resultSet = responsePayload.PrimaryQueryResult.RelevantResults.Table.Rows.map((result: any) => ({
            key:result.Cells[8].Value,
            //header:result.Cells[0].Value,
            headerMedia:result.Cells[2].Value,
            content:result.Cells[1].Value,
          }));        
      console.log(JSON.stringify(resultSet));
      setSearchResults(resultSet);
    }
  }
  */
    //Commenting this out to test MGT
    /*
    const MyMessage = (props: MgtTemplateProps) => {
        const message = props.dataContext as MicrosoftGraph.Message;
      
        const personRef = useRef<MgtPerson>();
      
        const handlePersonClick = () => {
          console.log(personRef.current);
        };
        return (
            <div>
              <b>Subject:</b>
              {message.subject}
              <div>
                <b>From:</b>
                <Person
                  ref={personRef}
                  onClick={handlePersonClick}
                  personQuery={message.from?.emailAddress?.address || ''}
                  fallbackDetails={{ mail: message.from?.emailAddress?.address, displayName: message.from?.emailAddress?.name }}
                  view={PersonViewType.oneline}
                ></Person>
              </div>
            </div>
          );
        };

      const SiteResult = (props: MgtTemplateProps) => {
          const site = props.dataContext as MicrosoftGraph.Site;

          return (
              <div>
                <Flex gap="gap.medium" padding="padding.medium" debug>
                <Flex.Item size="size.medium">
                  <div
                      style={{
                      position: 'relative',
                      }}
                  >
                      <Image
                      height={40}
                      width={40}
                      fluid
                      src="https://upload.wikimedia.org/wikipedia/commons/thumb/e/e1/Microsoft_Office_SharePoint_%282019%E2%80%93present%29.svg/2097px-Microsoft_Office_SharePoint_%282019%E2%80%93present%29.svg.png"
                      />
                  </div>
                  </Flex.Item>
                  <Flex.Item grow>
                  <Flex column gap="gap.small" vAlign="stretch">
                      <Flex space="between">
                      <Header as="h3" content={site.displayName} />
                      <Text as="pre" content={site.name} />
                      </Flex>

                      <Text content={site.webUrl} />

                      <Flex.Item push>
                      <Text as="pre" content="COPYRIGHT: Fluent UI." />
                      </Flex.Item>
                  </Flex>
                  </Flex.Item>                    
                </Flex>
              </div>
            );
          };

    const getRecentEmails = useCallback(async () => {
        if (!msGraphOboToken) { return; }
      
        const endpoint = `https://graph.microsoft.com/v1.0/me/messages?$select=receivedDateTime,subject&$orderby=receivedDateTime&$top=10`;
        const requestObject = {
          method: 'GET',
          headers: {
            "authorization": "bearer " + msGraphOboToken
          }
        };
      
        const response = await fetch(endpoint, requestObject);
        const responsePayload = await response.json();
      
        if (response.ok) {
          const recentMail = responsePayload.value.map((mail: any) => ({
            key: mail.id,
            header: mail.subject,
            headerMedia: mail.receivedDateTime
          }));
          setRecentMail(recentMail);
        }
      }, [msGraphOboToken]);

    const exchangeSsoTokenForOboToken = useCallback(async () => {
        const response = //await fetch(`https://azfun.ngrok.io/api/TeamsOBOHelper?ssoToken=${ssoToken}&tokenFor=msg`);
        await fetch(`http://localhost:5000/api/token?ssoToken=${ssoToken}&tokenFor=msg`);
        const responsePayload = await response.json();
        if (response.ok) {
          setMsGraphOboToken(responsePayload.access_token);
        } else {
          if (responsePayload!.error === "consent_required") {
            setError("consent_required");
          } else {
            setError("unknown SSO error");
          }
        }
      }, [ssoToken]); 


    useEffect(() => {
        // if the SSO token is defined...
        if (ssoToken && ssoToken.length > 0) {
          exchangeSsoTokenForOboToken();
        }
      }, [exchangeSsoTokenForOboToken, ssoToken]);
      

    useEffect(() => {
        getRecentEmails();
      }, [msGraphOboToken]);
     */

    /**
     * The render() method to create the UI of the tab
     */
    /*
    return (        
      <Pivot>        
        <PivotItem headerText="Sites Search Using SPO REST API">
        <div>
          {searchResults && <List items={searchResults} />}
        </div>
        </PivotItem>
        <PivotItem headerText="App">
        <App />
        </PivotItem>
    </Pivot>
    );*/

    return (
      <div>
          {
            ssoToken &&
            <Pivot aria-label="Basic Pivot Example">
                <PivotItem headerText="SPO REST API">
                    <SPOReusable idToken={ssoToken} />
                </PivotItem>
                <PivotItem headerText="MS Graph REST API">
                    <MSGReusable idToken={ssoToken} />
                </PivotItem>
                <PivotItem headerText="MS Graph Toolkit">
                    <Pivot>
                        <PivotItem headerText="Files">
                            <FileList></FileList> 
                        </PivotItem>
                        <PivotItem headerText="People">
                            <br/>
                            <PeoplePicker></PeoplePicker>
                        </PivotItem>
                        <PivotItem headerText="File Upload">
                            <FileList driveId="b!mKw3q1anF0C5DyDiqHKMr8iJr_oIRjlGl4854HhHtho07AdbOeaLT5rMH83yt89B" 
                        itemPath="/" enableFileUpload></FileList>
                        </PivotItem>
                        <PivotItem headerText="Sites Search Using MSGraph">
                            <Get resource="/sites?search=contoso" scopes={['Sites.Read.All']} maxPages={2}>
                                    <SiteResult template="value" />
                            </Get>
                        </PivotItem>
                    </Pivot>
                </PivotItem>           
                <PivotItem headerText="Adaptive Card">
                    <ShowAdaptiveCard />
                </PivotItem>        
            </Pivot>
          }
      </div>
    );

    return (        
      <Pivot>        
        <PivotItem headerText="Sites Search Using SPO REST API">
        <div>
          {ssoToken && <ReusableApp idToken={ssoToken} />}
        </div>
        </PivotItem>
        <PivotItem headerText="App">
        <App />
        </PivotItem>
    </Pivot>
    );
};

/*
return (
<Provider theme={theme}>
            <Flex fill={true} column styles={{
                padding: ".8rem 0 .8rem .5rem"
            }}>
                <Flex.Item>
                    <Header content={`Hello ${name}`}  />
                </Flex.Item>
                <Flex.Item>
                    <div>
                        {error && <div><Text content={`An SSO error occurred ${error}`} /></div>}
                        <div>                  
                          <Login />
                        </div>
                        <div>
                        <Person
                            personQuery="me"
                            view={PersonViewType.twolines}
                            personCardInteraction={PersonCardInteraction.hover}
                            showPresence={true}
                        />
                        </div>
                        <div>
                            <PeoplePicker></PeoplePicker>
                        </div>
                        <div>
                        <ul className="breadcrumb" id="nav">
                            <li><a id="home">Files</a></li>
                        </ul>
                        <FileList
                                driveId="b!mKw3q1anF0C5DyDiqHKMr8iJr_oIRjlGl4854HhHtho07AdbOeaLT5rMH83yt89B" 
                                itemPath="/" enableFileUpload 
                                />

                        </div>
                        <div>
                        <Get resource="/sites?search=contoso" scopes={['Sites.Read.All']} maxPages={2}>
                            <SiteResult template="value" />
                        </Get>
                        </div>                        

                        
                                                <div>
                        <Get resource="/me/messages" scopes={['mail.read']} maxPages={2}>
                            <MyMessage template="value" />
                        </Get>
                        </div>
                        <div>
                            {recentMail && <div><h3>Your recent emails:</h3><List items={recentMail} /></div>}
                        </div>
                        
                        <div>
                            <Agenda  />
                        </div> 
                        <div>
                                                
                        <FileList
                            siteId="m365x229910.sharepoint.com,ab37ac98-a756-4017-b90f-20e2a8728caf,faaf89c8-4608-4639-978f-39e07847b61a" 
                            itemPath="/" enableFileUpload />
                        </div>
                        
                    </div>
                    
                </Flex.Item>
                <Flex.Item styles={{
                    padding: ".8rem 0 .8rem .5rem"
                }}>
                    <Text size="smaller" content="(C) Copyright Contoso" />
                </Flex.Item>
            </Flex>
        </Provider>
);
*/