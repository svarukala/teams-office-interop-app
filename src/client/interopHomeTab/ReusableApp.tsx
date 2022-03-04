import * as React from 'react';
import { useState, useEffect } from 'react';
import * as msal from "@azure/msal-browser";
//import { List } from 'office-ui-fabric-react';
import { List } from "@fluentui/react-northstar";


function ReusableApp(props) {
  // Declare a new state variable, which we'll call "count"
  const [totalresults, setTotalresults] = useState(0);
  const [ssoToken, setSsoToken] = useState<string>();
  const [msGraphOboToken, setMsGraphOboToken] = useState<string>();
  const [spoOboToken, setSPOOboToken] = useState<string>();
  const [searchResults, setSearchResults] = useState<any[]>();
  const [error, setError] = useState<string>();

  useEffect(() => {
    // if the SSO token is defined...
    if (props.idToken) {
      setSsoToken(props.idToken);
    }
  }, []);    

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

  return (
    <div>
        {error && "Error: " + error}
        {ssoToken && "ID Token: " + ssoToken}
        {searchResults && <List items={searchResults} />}
    </div>
  );
}

export default ReusableApp;


/*
<p>You clicked {count} times</p>
<button onClick={() => setCount(count + 1)}>
  Click me
</button>
*/