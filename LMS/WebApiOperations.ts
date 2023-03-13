/*global Xrm */
/*eslint no-undef: "error"*/
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
  selectedKey: string | number,
  primaryEntityType: string,
  relatedEntityType: string,
  relationshipName: string,
  primaryEntityId: string
) => {
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
        console.log(this.statusText);
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
  removedKey: string | number,
  primaryEntityType: string,
  relationshipName: string,
  primaryEntityId: string
) => {
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
        console.log(this.statusText);
      }
    }
  };
  req.send();
};
