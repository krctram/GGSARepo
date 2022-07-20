import * as React from "react";
import { useState, useEffect } from "react";
import "@pnp/sp/webs";
import { Web } from "@pnp/sp/webs";
import "@pnp/sp/site-users/web";
import "@pnp/sp/site-groups/web";
import AnnualPlan from "./AnnualPlan";
import DeliveryPlan from "./DeliveryPlan";
import ProductionBoard from "./ProductionBoard";
import DocumentReview from "./DocumentReview";
import ActivityPlan from "./ActivityPlan";
import ActivityDeliveryPlan from "./ActivityDeliveryPlan";
import ActivityTemplate from "./ActivityTemplate";

/*development URL */
const webURL = "https://ggsaus.sharepoint.com/sites/Intranet_Test";
const WeblistURL = "Annual Plan Test";

/*Production URL */
//const webURL = "https://ggsaus.sharepoint.com";
//const WeblistURL = "Annual Plan";

const MainComponent = (props: any) => {
  const _webURL = Web(webURL);

  const [pageSwitch, setPageSwitch] = useState("AnnualPlan");
  const [annualPlanID, setAnnualPlanID] = useState();
  const [adminStatus, setAdminStatus] = useState(true);
  const [pageType, setpageType] = useState();

  //let PageName;
  const handleclick = (page: string, AP_ID: any, type: any) => {
    // PageName = page;
    setPageSwitch(page);
    setAnnualPlanID(AP_ID);
    setpageType(type);
  };
  const getAdmins = () => {
    _webURL.siteGroups
      .getByName("Innovation Hub Admin")
      .users.get()
      .then((users) => {
        props.context.web
          .currentUser()
          .then((user) => {
            let tempUser = users.filter((_user) => {
              return _user.Email == user.Email;
            });
            if (tempUser.length > 0) {
              setAdminStatus(true);
            }
          })
          .catch((error) => {
            alert(error);
          });
      })
      .catch((error) => {
        alert(error);
      });
  };
  useEffect(() => {
    //getAdmins();
    const urlParams = new URLSearchParams(window.location.search);
    const pageName = urlParams.get("Page");

    if (pageName == "AP") {
      setPageSwitch("AnnualPlan");
    } else if (pageName == "PB") {
      setPageSwitch("ProductionBoard");
    } else if (pageName == "DR") {
      setPageSwitch("DocumentReview");
    } else if (pageName == "ATP") {
      setPageSwitch("ActivityPlan");
    }
    // else {
    //   setPageSwitch(PageName);
    // }
  }, []);
  return (
    <>
      {pageSwitch == "AnnualPlan" ? (
        <AnnualPlan
          context={props.context}
          spcontext={props.spcontext}
          graphContent={props.graphContent}
          URL={webURL}
          WeblistURL={WeblistURL}
          isAdmin={adminStatus}
          handleclick={handleclick}
          pageType={pageType}
        />
      ) : pageSwitch == "DeliveryPlan" ? (
        <DeliveryPlan
          context={props.context}
          spcontext={props.spcontext}
          graphContent={props.graphContent}
          URL={webURL}
          WeblistURL={WeblistURL}
          handleclick={handleclick}
          AnnualPlanId={annualPlanID}
          pageType={pageType}
        />
      ) : pageSwitch == "ProductionBoard" ? (
        <ProductionBoard
          context={props.context}
          spcontext={props.spcontext}
          graphContent={props.graphContent}
          URL={webURL}
          WeblistURL={WeblistURL}
          handleclick={handleclick}
          AnnualPlanId={annualPlanID}
          pageType={pageType}
        />
      ) : pageSwitch == "DocumentReview" ? (
        <DocumentReview
          context={props.context}
          spcontext={props.spcontext}
          graphContent={props.graphContent}
          URL={webURL}
          handleclick={handleclick}
          pageType={pageType}
        />
      ) : pageSwitch == "ActivityPlan" ? (
        <ActivityPlan
          context={props.context}
          spcontext={props.spcontext}
          graphContent={props.graphContent}
          URL={webURL}
          handleclick={handleclick}
          pageType={pageType}
        />
      ) : pageSwitch == "ActivityDeliveryPlan" ? (
        <ActivityDeliveryPlan
          context={props.context}
          spcontext={props.spcontext}
          graphContent={props.graphContent}
          URL={webURL}
          handleclick={handleclick}
          pageType={pageType}
        />
      ) : pageSwitch == "ActivityTemplate" ? (
        <ActivityTemplate
          context={props.context}
          spcontext={props.spcontext}
          graphContent={props.graphContent}
          URL={webURL}
          handleclick={handleclick}
          pageType={pageType}
        />
      ) : (
        ""
      )}
    </>
  );
};

export default MainComponent;
