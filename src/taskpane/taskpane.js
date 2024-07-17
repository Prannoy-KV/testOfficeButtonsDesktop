/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.PowerPoint) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("enableSignout").onclick = enableSignout;
    document.getElementById("disableSignout").onclick = disableSignout;
  }

  // setTimeout(()=>{
  //   test();
  // },3000);
  
});

async function test(){
  const buttonsToBeUpdated = [
    { id: "Starlight.LogoutButton", enabled: true }   
  ];
  const parentTab = { id: "Starlight.Tab1", groups: [ { id: "Starlight.Tab1.Group1", controls: buttonsToBeUpdated}] };
  const ribbonUpdater= { tabs: [parentTab] };
  try {
    await Office.ribbon.requestUpdate(ribbonUpdater);
  } catch (error) {
    console.log(error);
  }
}
export async function disableSignout() {
  /**
   * Insert your PowerPoint code here
   */
  // const options = { coercionType: Office.CoercionType.Text };

  // await Office.context.document.setSelectedDataAsync(" ", options);
  // await Office.context.document.setSelectedDataAsync("Hello World!", options);
  const buttonsToBeUpdated = [
    { id: "Starlight.LoginButton", enabled: true },
    { id: "Starlight.LogoutButton", enabled: false }   
  ];
  const parentTab = { id: "Starlight.Tab1", groups: [ { id: "Starlight.Tab1.Group1", controls: buttonsToBeUpdated}] };
  const ribbonUpdater= { tabs: [parentTab] };
  try {
    await Office.ribbon.requestUpdate(ribbonUpdater);
  } catch (error) {
    console.log(error);
  }

}

export async function enableSignout() {
  /**
   * Insert your PowerPoint code here
   */
  // const options = { coercionType: Office.CoercionType.Text };

  // await Office.context.document.setSelectedDataAsync(" ", options);
  // await Office.context.document.setSelectedDataAsync("Hello World!", options);
  const buttonsToBeUpdated = [
    { id: "Starlight.LoginButton", enabled: false },
    { id: "Starlight.LogoutButton", enabled: true }   
  ];
  const parentTab = { id: "Starlight.Tab1", groups: [ { id: "Starlight.Tab1.Group1", controls: buttonsToBeUpdated}] };
  const ribbonUpdater= { tabs: [parentTab] };
  try {
    await Office.ribbon.requestUpdate(ribbonUpdater);
  } catch (error) {
    console.log(error);
  }

}