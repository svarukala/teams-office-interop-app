import * as React from "react";
import { Provider, Flex, Image, Text, Button, Header, List, ItemLayout } from "@fluentui/react-northstar";
import { useState, useEffect, useCallback, useRef } from "react";
import { useTeams } from "msteams-react-base-component";
import { app, authentication } from "@microsoft/teams-js";
import jwtDecode from "jwt-decode";
import {Providers, ProviderState, LoginType} from '@microsoft/mgt-element';
import * as MicrosoftTeams from "@microsoft/teams-js";
import {TeamsMsal2Provider, HttpMethod} from '@microsoft/mgt-teams-msal2-provider';
import { Msal2Provider } from '@microsoft/mgt-msal2-provider';
import { Login, Person, FileList, Agenda, PersonViewType, PeoplePicker, PersonCardInteraction, MgtTemplateProps, Get } from '@microsoft/mgt-react';
import { MgtPerson } from '@microsoft/mgt-components';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import App from './App';
/**
 * Implementation of the Interop Home content page
 */
 
export const InteropHomeTab = () => {

    const [{ inTeams, theme, context }] = useTeams();
    const [entityId, setEntityId] = useState<string | undefined>();
    const [name, setName] = useState<string>();
    const [error, setError] = useState<string>();
    const [ssoToken, setSsoToken] = useState<string>();
    const [msGraphOboToken, setMsGraphOboToken] = useState<string>();
    const [spoOboToken, setSPOOboToken] = useState<string>();
    const [recentMail, setRecentMail] = useState<any[]>();

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
        }
    }, [inTeams]);

    useEffect(() => {
        if (context) {
            setEntityId(context.page.id);
        }
    }, [context]);

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
    return (
        <App />
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