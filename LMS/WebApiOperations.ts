/*global Xrm */
/*eslint no-undef: "error"*/

import { IInputs } from "./generated/ManifestTypes";

//@ts-ignore
const url: string = Xrm.Utility.getGlobalContext().getClientUrl();

/**
 * Associate selected record
 * @param selectedKey
 * @param primaryEntityType
 * @param relatedEntityType
 * @param relationshipName
 * @param primaryEntityId
 */
export const Associate = (
  context: ComponentFramework.Context<IInputs>,
  selectedKey: string | number,
  primaryEntityType: string,
  relatedEntityType: string,
  relationshipName: string,
  primaryEntityId: string
) => {
  let error: any =
    "Issue while associating the record. Kindly check the configuration.";
  var association = {
    "@odata.id": url + `/api/data/v9.1/${relatedEntityType}s(${selectedKey})`,
  };
  var req = new XMLHttpRequest();
  req.open(
    "POST",
    url +
      `/api/data/v9.1/${primaryEntityType}s(${primaryEntityId})/${relationshipName}/$ref`,
    true
  );
  req.setRequestHeader("Accept", "application/json");
  req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
  req.setRequestHeader("OData-MaxVersion", "4.0");
  req.setRequestHeader("OData-Version", "4.0");
  req.onreadystatechange = function () {
    if (this.readyState === 4) {
      req.onreadystatechange = null;
      if (this.status === 204 || this.status === 1223) {
        console.log("Associated");
      } else {
        context.navigation.openAlertDialog(error);
      }
    }
  };
  req.send(JSON.stringify(association));
};

/**
 * Disassociate cancelled/unselected record
 * @param removedKey
 * @param primaryEntityType
 * @param relationshipName
 * @param primaryEntityId
 */
export const DisAssociate = (
  context: ComponentFramework.Context<IInputs>,
  removedKey: string | number,
  primaryEntityType: string,
  relationshipName: string,
  primaryEntityId: string
) => {
  let error: any =
    "Issue while disassociating the record. Kindly check the configuration.";
  var req = new XMLHttpRequest();
  req.open(
    "DELETE",
    url +
      `/api/data/v9.1/${primaryEntityType}s(${primaryEntityId})/${relationshipName}(${removedKey})/$ref`,
    true
  );
  req.setRequestHeader("Accept", "application/json");
  req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
  req.setRequestHeader("OData-MaxVersion", "4.0");
  req.setRequestHeader("OData-Version", "4.0");
  req.onreadystatechange = function () {
    if (this.readyState === 4) {
      req.onreadystatechange = null;
      if (this.status === 204 || this.status === 1223) {
        console.log("Disassociated");
      } else {
        context.navigation.openAlertDialog(error);
      }
    }
  };
  req.send();
};
