import * as React from "react";
import { useState, useEffect } from "react";
import * as moment from "moment";
import { Web } from "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/fields";
import "@pnp/sp/site-users/web";
import {
  DetailsList,
  DetailsListLayoutMode,
  IDetailsListStyles,
  SelectionMode,
  Icon,
  Label,
  ILabelStyles,
  Dropdown,
  IDropdownStyles,
  NormalPeoplePicker,
  Persona,
  PersonaPresence,
  PersonaSize,
  Modal,
  IModalStyles,
  DatePicker,
  IDatePickerStyles,
  MessageBar,
  MessageBarType,
  Spinner,
  SpinnerSize,
  TooltipHost,
  TooltipOverflowMode,
} from "@fluentui/react";
import "../ExternalRef/styleSheets/activityStyles.css";
import { mergeStyles, mergeStyleSets } from "@fluentui/react/lib/Styling";
import styles from "./InnovationHubIntranet.module.scss";
import CustomLoader from "./CustomLoader";

const ActivityPlan = (props: any) => {
  // Variable-Declaration-Section Starts
  const sharepointWeb = Web(props.URL);
  const ListName = "Activity Plan";

  const atpAllitems = [];
  const atpAllPeoples = [];
  const atpColumns = [
    {
      key: "Lesson",
      name: "Lesson",
      fieldName: "Lesson",
      minWidth: 300,
      maxWidth: 300,
    },
    {
      key: "Project",
      name: "Project",
      fieldName: "Project",
      minWidth: 300,
      maxWidth: 300,
    },
    {
      key: "StartDate",
      name: "StartDate",
      fieldName: "StartDate",
      minWidth: 100,
      maxWidth: 100,
      onRender: (item) => moment(item.StartDate).format("DD/MM/YYYY"),
    },
    {
      key: "EndDate",
      name: "EndDate",
      fieldName: "EndDate",
      minWidth: 100,
      maxWidth: 100,
      onRender: (item) => moment(item.EndDate).format("DD/MM/YYYY"),
    },
    {
      key: "Developer",
      name: "Developer",
      fieldName: "Developer",
      minWidth: 350,
      maxWidth: 350,

      onRender: (item) => (
        <div style={{ display: "flex" }}>
          <div
            // title={item.DeveloperDetails.name}
            style={{
              marginTop: "-6px",
            }}
          >
            <Persona
              size={PersonaSize.size32}
              presence={PersonaPresence.none}
              imageUrl={
                "/_layouts/15/userphoto.aspx?size=S&username=" +
                `${item.DeveloperDetails.email}`
              }
            />
          </div>
          <div>
            <Label>{item.DeveloperDetails.name}</Label>
          </div>
        </div>
      ),
    },
    {
      key: "Action",
      name: "Action",
      fieldName: "Action",
      minWidth: 70,
      maxWidth: 70,

      onRender: (item) => (
        <>
          <Icon
            iconName="Edit"
            className={atpIconStyleClass.edit}
            onClick={() => {
              console.log(`Edit - ${item.ID}`);
            }}
          />
          <Icon
            iconName="Delete"
            className={atpIconStyleClass.delete}
            onClick={() => {
              console.log(`Delete - ${item.ID}`);
            }}
          />
        </>
      ),
    },
    {
      key: "Link",
      name: "Link",
      fieldName: "Link",
      minWidth: 70,
      maxWidth: 70,

      onRender: (item) => (
        <>
          <Icon
            iconName="Link12"
            className={atpIconStyleClass.link}
            onClick={() => {
              props.handleclick("ActivityDeliveryPlan");
            }}
          />
        </>
      ),
    },
  ];
  const atpDrpDwnOptns = {
    lessonOptns: [{ key: "All", text: "All" }],
    projectOptns: [{ key: "All", text: "All" }],
  };
  const atpFilterKeys = {
    lesson: "All",
    project: "All",
    start: null,
    end: null,
    dev: null,
  };
  // Variable-Declaration-Section Ends
  // Styles-Section Starts
  const atpDetailsListStyles: Partial<IDetailsListStyles> = {
    root: {},
    headerWrapper: {},
    contentWrapper: {},
  };
  const atpLabelStyles: Partial<ILabelStyles> = {
    root: {
      width: 150,
      marginRight: 10,
      fontSize: "13px",
      color: "#323130",
    },
  };
  const atpDropdownStyles: Partial<IDropdownStyles> = {
    root: {
      width: 165,
      marginRight: 15,
      backgroundColor: "#F5F5F7",
    },
    title: {
      backgroundColor: "#F5F5F7",
      fontSize: 12,
      color: "#7C7C7C",
      border: "1px solid #E8E8EA",
      borderRadius: "4px",
    },
    dropdownItem: {
      backgroundColor: "#F5F5F7",
      fontSize: 12,
    },
    dropdownItemSelected: {
      backgroundColor: "#F5F5F7",
      fontSize: 12,
    },
    caretDown: { fontSize: 14, color: "#000" },
  };
  const atpModalBoxDatePickerStyles: Partial<IDatePickerStyles> = {
    root: {
      width: 165,
      marginRight: 15,
      marginTop: 5,
      borderRadius: "4px",
    },
    icon: {
      fontSize: "17px",
      color: "#000",
      fontWeight: "bold",
    },
  };
  const atpModalStyles: Partial<IModalStyles> = {
    root: { borderRadius: "none" },
    main: {
      width: 500,
      margin: 10,
      padding: "20px 10px",
      display: "flex",
      flexDirection: "column",
      alignItems: "center",
      justifyContent: "center",
    },
  };
  const atpIconStyleClass = mergeStyleSets({
    link: [
      {
        fontSize: 17,
        height: 14,
        width: 17,
        color: "#2392B2",
        cursor: "pointer",
      },
    ],
    edit: [
      {
        fontSize: 17,
        height: 14,
        width: 17,
        color: "#2392B2",
        cursor: "pointer",
      },
    ],
    delete: [
      {
        fontSize: 17,
        height: 14,
        width: 17,
        marginLeft: 10,
        color: "#CB1E06",
        cursor: "pointer",
      },
    ],
    linkPB: [
      {
        fontSize: 18,
        height: 16,
        width: 19,
        color: "#fff",
        backgroundColor: "#038387",
        cursor: "pointer",
        padding: 8,
        borderRadius: 3,
        marginLeft: 10,
        ":hover": {
          backgroundColor: "#025d60",
        },
      },
    ],
    refresh: [
      {
        fontSize: 18,
        height: 16,
        width: 19,
        color: "#fff",
        backgroundColor: "#038387",
        cursor: "pointer",
        padding: 8,
        borderRadius: 3,
        marginLeft: 10,
        marginTop: 27,
        ":hover": {
          backgroundColor: "#025d60",
        },
      },
    ],
  });
  const generalStyles = mergeStyleSets({
    titleLabel: {
      color: "#2392B2 !important",
      fontWeight: "500",
      fontSize: "17px",
    },
    inputLabel: {
      color: "#2392B2 !important",
      display: "block",
      fontWeight: "500",
      margin: "5px 0",
    },
    inputValue: {
      color: "#000",
      fontWeight: "500",
      fontSize: "13px",
    },
    inputField: {
      margin: "10px 0",
    },
  });
  // Styles-Section Ends
  // States-Declaration Starts
  const [atpReRender, setAtpReRender] = useState(true);
  const [currentUser, setCurrentUser] = useState({});
  const [atpMasterData, setAtpMasterData] = useState(atpAllitems);
  const [atpData, setAtpData] = useState(atpAllitems);
  const [atpPeopleList, setAtpPeopleList] = useState(atpAllPeoples);
  const [atpDropDownOptions, setAtpDropDownOptions] = useState(atpDrpDwnOptns);
  const [atpFilters, setAtpFilters] = useState(atpFilterKeys);
  const [atpAddPlannerPopup, setAtpAddPlannerPopup] = useState(false);
  const [atpLoader, setAtpLoader] = useState("noLoader");
  const [atpPopup, setAtpPopup] = useState("close");
  // States-Declaration Ends
  //Function-Section Starts
  const getAllUsers = () => {
    sharepointWeb
      .siteUsers()
      .then((_allUsers) => {
        _allUsers.forEach((user) => {
          atpAllPeoples.push({
            key: 1,
            imageUrl:
              `/_layouts/15/userphoto.aspx?size=S&accountname=` +
              `${user.Email}`,
            text: user.Title,
            ID: user.Id,
            secondaryText: user.Email,
            isValid: true,
          });
        });
        setAtpPeopleList(atpAllPeoples);
      })
      .catch(atpErrorFunction);
  };
  const atpGetData = () => {
    sharepointWeb.lists
      .getByTitle(ListName)
      .items.select("*", "Developer/Title", "Developer/Id", "Developer/EMail")
      .expand("Developer")
      .get()
      .then((items) => {
        console.log(items);
        items.forEach((item) => {
          atpAllitems.push({
            ID: item.Id ? item.Id : "",
            Lesson: item.Title ? item.Title : "",
            Project: item.Project ? item.Project : "",
            StartDate: item.StartDate ? item.StartDate : null,
            EndDate: item.EndDate ? item.EndDate : null,
            Developer: item.DeveloperId ? item.Developer.Title : "",
            DeveloperDetails: item.DeveloperId
              ? {
                  name: item.Developer.Title,
                  id: item.Developer.Id,
                  email: item.Developer.EMail,
                  userDetails: atpPeopleList.filter((people) => {
                    return people.ID == item.Developer.Id;
                  })[0],
                }
              : "",
          });
        });

        // console.log(atpAllitems);

        setAtpMasterData([...atpAllitems]);
        setAtpData([...atpAllitems]);
        setAtpLoader("noLoader");
      })
      .catch(atpErrorFunction);
  };
  const dateFormater = (date: Date): string => {
    return date ? moment(date).format("DD/MM/YYYY") : "";
  };
  const SubmitPopup = () => (
    <MessageBar messageBarType={MessageBarType.success} isMultiline={false}>
      Document review has been successfully submitted !!!
    </MessageBar>
  );
  const atpErrorFunction = (error: any) => {
    console.log(error);
    setAtpPopup("Error");
    setTimeout(() => {
      setAtpPopup("Close");
    }, 2000);
  };
  //Function-Section Ends
  useEffect(() => {
    setAtpLoader("startUpLoader");
    getAllUsers();
    atpGetData();
  }, [atpReRender]);
  return (
    <>
      <div style={{ padding: "5px 10px" }}>
        {atpLoader == "startUpLoader" ? <CustomLoader /> : null}
        {/* Header-Section Starts */}
        <div className={styles.atpHeaderSection}>
          {/* Popup-Section Starts */}
          <div>{atpPopup == "Submit" ? SubmitPopup() : ""}</div>
          {/* Popup-Section Ends */}
          <div className={styles.atpHeader}>Activity plan</div>
          <div style={{ display: "flex", justifyContent: "space-between" }}>
            <div>
              <button
                className={styles.atpAddBtn}
                onClick={() => {
                  props.handleclick("ActivityTemplate", null, "AT");
                }}
              >
                Add template
              </button>
            </div>
            <div style={{ display: "flex" }}>
              <button
                className={styles.atpAddBtn}
                onClick={() => {
                  setAtpAddPlannerPopup(true);
                }}
              >
                Add Activity planner
              </button>
              <Icon
                iconName="Link12"
                className={atpIconStyleClass.linkPB}
                onClick={() => {}}
              />
            </div>
          </div>
        </div>
        {/* Header-Section Ends */}
        {/* Filter-Section Starts */}
        <div>
          <div
            style={{
              display: "flex",
              alignItems: "center",
              justifyContent: "flex-start",
              paddingBottom: "10px",
            }}
          >
            <div>
              <Label styles={atpLabelStyles}>Lesson</Label>
              <Dropdown
                placeholder="Select an option"
                styles={atpDropdownStyles}
                options={atpDropDownOptions.lessonOptns}
                dropdownWidth={"auto"}
                onChange={(e, option: any) => {}}
                selectedKey={atpFilters.lesson}
              />
            </div>
            <div>
              <Label styles={atpLabelStyles}>Project</Label>
              <Dropdown
                placeholder="Select an option"
                styles={atpDropdownStyles}
                options={atpDropDownOptions.projectOptns}
                dropdownWidth={"auto"}
                onChange={(e, option: any) => {}}
                selectedKey={atpFilters.project}
              />
            </div>
            <div>
              <Label styles={atpLabelStyles}>Start Date</Label>
              <DatePicker
                formatDate={dateFormater}
                value={new Date()}
                styles={atpModalBoxDatePickerStyles}
                onSelectDate={(value: any) => {}}
              />
            </div>
            <div>
              <Icon
                iconName="Refresh"
                className={atpIconStyleClass.refresh}
                onClick={() => {}}
              />
            </div>
          </div>
        </div>
        {/* Filter-Section Ends */}
        {/* Body-Section Starts */}
        <div>
          {/* DetailList-Section Starts */}
          {atpData.length > 0 ? (
            <div>
              <DetailsList
                items={atpData}
                columns={atpColumns}
                styles={atpDetailsListStyles}
                setKey="set"
                layoutMode={DetailsListLayoutMode.justified}
                selectionMode={SelectionMode.none}
              />
            </div>
          ) : (
            <Label
              style={{
                paddingLeft: 745,
                paddingTop: 40,
              }}
              className={generalStyles.inputLabel}
            >
              No Data Found !!!
            </Label>
          )}
          {/* DetailList-Section Ends */}
        </div>
        {/* Body-Section Ends */}
        {/* Modal-Section Starts */}
        {atpAddPlannerPopup ? (
          <Modal
            isOpen={atpAddPlannerPopup}
            isBlocking={true}
            styles={atpModalStyles}
          >
            <div>
              <Label className={styles.drPopupLabel}>Confirmation</Label>
              <div className={styles.drPopupDescription}>
                This will cancel the request and remove from the persons review
                log. Enter your reason to cancel.
              </div>
              <div className={styles.drPopupDescription}>
                Do you wish to proceed?
              </div>
            </div>
          </Modal>
        ) : (
          ""
        )}
        {/* Modal-Section Ends */}
      </div>
    </>
  );
};

export default ActivityPlan;
