import * as React from 'react';
import { useState, useEffect } from 'react';
import * as msal from "@azure/msal-browser";
import { List } from 'office-ui-fabric-react';

let currentAccount: msal.AccountInfo = null;

const msalConfig = {  
    auth: {  
      clientId: 'c613e0d1-161d-4ea0-9db4-0f11eeabc2fd',  
      redirectUri: 'https://m365x229910.sharepoint.com/_layouts/15/workbench.aspx'  
    }  
  };
  
const msalInstance = new msal.PublicClientApplication(msalConfig);

const tokenrequest: msal.SilentRequest = {
    scopes: ["https://m365x229910.sharepoint.com/AllSites.Read", "https://m365x229910.sharepoint.com/AllSites.Manage"],
    account: currentAccount,
    };



function SPOSearch(props) {
  // Declare a new state variable, which we'll call "count"
  const [count, setCount] = useState(0);
  const [resultitems, setResultitems] = useState([]);
  const [totalresults, setTotalresults] = useState(0);

  const setCurrentAccount = (): void => {
    const currentAccounts: msal.AccountInfo[] = msalInstance.getAllAccounts();
    if (currentAccounts === null || currentAccounts.length == 0) {
        tokenrequest.account = msalInstance.getAccountByUsername(
        this.context.pageContext.user.loginName
      );
    } else if (currentAccounts.length > 1) {
      console.warn("Multiple accounts detected.");
      currentAccount = msalInstance.getAccountByUsername(
        this.context.pageContext.user.loginName
      );
    } else if (currentAccounts.length === 1) {
      currentAccount = currentAccounts[0];
    }
    tokenrequest.account = currentAccount;
  }; 

    useEffect(() => {
        setCurrentAccount();
        msalInstance.acquireTokenSilent(tokenrequest).then((val) => {  
        let headers = new Headers();  
        let bearer = "Bearer " + val.accessToken;  
        console.info("BEARER TOKEN: "+ val.accessToken);
        headers.append("Authorization", bearer); 
        headers.append("Accept", "application/json;odata=verbose");
        headers.append("Content-Type", "application/json;odata=verbose");
        let options = {  
            method: "GET",  
            headers: headers  
        };  
        const res = fetch("https://m365x229910.sharepoint.com/sites/DevDemo/_api/search/query?querytext='*'", options);
        res.then(resp => {  
                resp.json().then((data) => {  
                    console.log(data);
                    //var jsonObject = JSON.parse(data);
                    
                    console.log("Total results: "+ data.d.query.PrimaryQueryResult.RelevantResults.TotalRows);
                    
                    setResultitems(data.d.query.PrimaryQueryResult.RelevantResults.Table.Rows.results);
                    setTotalresults(data.d.query.PrimaryQueryResult.RelevantResults.TotalRows);
                });
            });  
        }).catch((errorinternal) => {  
            console.log(errorinternal);  
        });  
    }, []);

  return (
    <div>
        <h2>{totalresults && "Total Results: "+ totalresults}</h2>
      <div><pre>{resultitems && JSON.stringify(resultitems, null, 2) }</pre></div>
    </div>
  );
}

export default SPOSearch;


/*
<p>You clicked {count} times</p>
<button onClick={() => setCount(count + 1)}>
  Click me
</button>
*/