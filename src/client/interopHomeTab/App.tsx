import { Agenda, Login, FileList, Get, MgtTemplateProps, PeoplePicker, Person, ViewType} from '@microsoft/mgt-react';
import { Grid, Card, CardHeader, CardBody, Flex, Text, Button, Header, Avatar, ItemLayout } from "@fluentui/react-northstar";
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import * as React from "react";
import { useState, useEffect } from 'react';
import { ImageFit, Pivot, PivotItem, Image } from 'office-ui-fabric-react';
import { Providers, ProviderState } from '@microsoft/mgt-element';


function useIsSignedIn(): [boolean] {
  const [isSignedIn, setIsSignedIn] = useState(false);

  useEffect(() => {
    const updateState = () => {
      const provider = Providers.globalProvider;
      setIsSignedIn(provider && provider.state === ProviderState.SignedIn);
    };

    Providers.onProviderUpdated(updateState);
    updateState();

    return () => {
      Providers.removeProviderUpdatedListener(updateState);
    }
  }, []);

  return [isSignedIn];
}


function App() {
    const [isSignedIn] = useIsSignedIn();
  
    return (
      <div className="App">
        <header>
          <Login />
        </header>
        {
            isSignedIn &&
          
            <Pivot aria-label="Basic Pivot Example">
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
                <PivotItem headerText="Sites Search">
                <Get resource="/sites?search=contoso" scopes={['Sites.Read.All']} maxPages={2}>
                      <SiteResult template="value" />
              </Get>
                </PivotItem>
            </Pivot>
        }
        </div>
    );
}  
{/*
function App() {
  const [isSignedIn] = useIsSignedIn();

  return (
    <div className="App">
      <header>
        <Login />
      </header>
      <div>
            <PeoplePicker></PeoplePicker>
        </div>
        <div>
        <ul className="breadcrumb" id="nav">
            <li><a id="home">Files</a></li>
        </ul>
        </div>
      <div>
        {isSignedIn &&
          <FileList driveId="b!mKw3q1anF0C5DyDiqHKMr8iJr_oIRjlGl4854HhHtho07AdbOeaLT5rMH83yt89B" 
          itemPath="/" enableFileUpload></FileList>
          }
      </div> 
      <div>
      {isSignedIn &&
            <Get resource="/sites?search=contoso" scopes={['Sites.Read.All']} maxPages={2}>
                    <SiteResult template="value" />
            </Get>
        }
      </div>     
    </div>
  );
}
*/}

const SiteResult = (props: MgtTemplateProps) => {
    const site = props.dataContext as MicrosoftGraph.Site;

    return (
        <div className="ms-ListBasicExample-itemCell">
        <Image
          className="ms-ListBasicExample-itemImage"
          src="https://upload.wikimedia.org/wikipedia/commons/thumb/e/e1/Microsoft_Office_SharePoint_%282019%E2%80%93present%29.svg/2097px-Microsoft_Office_SharePoint_%282019%E2%80%93present%29.svg.png"
          width={50}
          height={50}
          imageFit={ImageFit.cover}
        />
          <div className='site'>
              <div className="title">
                        <a href={site.webUrl??""} target="_blank" rel="noreferrer">
                            <h3>{site.displayName}</h3>
                        </a>
                        <span className="date">
                            {new Date(site.createdDateTime??"").toLocaleDateString()}
                        </span>
                    </div>
        </div>
        </div>
      );
    };

export default App;
