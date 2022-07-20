import * as React from "react";
import { useState, useEffect } from "react";
// const BackIconImg = require("../ExternalRef/assets/backIcon.png");
// const upArrowImg = require("../ExternalRef/assets/upArrow.png");
// const downArrowImg = require("../ExternalRef/assets/down arrow.png");
import * as moment from "moment";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { Web } from "@pnp/sp/webs";
import {
  DetailsList,
  DetailsListLayoutMode,
  SelectionMode,
  Icon,
  Label,
  Dropdown,
  IDropdownStyles,
  Persona,
  PersonaPresence,
  PersonaSize,
  Modal,
  DatePicker,
  NormalPeoplePicker,
  PrimaryButton,
  ChoiceGroup,
  TextField,
  ITextFieldStyles,
  Checkbox,
  MessageBar,
  MessageBarType,
  Spinner,
  SpinnerSize,
  TooltipHost,
  TooltipOverflowMode,
} from "@fluentui/react";
import { FontIcon, IIconProps } from "@fluentui/react/lib/Icon";
import '../ExternalRef/styleSheets/ProdStyles.css';
import { mergeStyles, mergeStyleSets } from "@fluentui/react/lib/Styling";
import styles from "./InnovationHubIntranet.module.scss";
import CustomLoader from "./CustomLoader";
import { IDetailsListStyles } from "office-ui-fabric-react";

const saveIcon: IIconProps = { iconName: "Save" };
const editIcon: IIconProps = { iconName: "Edit" };
const cancelIcon: IIconProps = { iconName: "Cancel" };
function DeliveryPlan(props: any) {
  const sharepointWeb = Web(props.URL);
  const ListNameURL = props.WeblistURL;
  let Ap_AnnualPlanId = props.AnnualPlanId;
  let Dp_Year = moment().year();
  let Dp_WeekNumber = moment().isoWeek();
  let loggeduseremail = props.spcontext.pageContext.user.email;

  // Items in Detail List
  let _dpAllitems = [];
  const allPeoples = [];
  const _dpColumns = [
    {
      key: "Column1",
      name: "Source",
      fieldName: "Source",
      minWidth: 50,
      maxWidth: 50,
    },
    {
      key: "Column2",
      name: "Activity",
      fieldName: "Title",
      minWidth: 250,
      maxWidth: 250,
      onRender: (item) => (
        <>
          <TooltipHost
            id={item.ID}
            content={item.Title}
            overflowMode={TooltipOverflowMode.Parent}
          >
            <span aria-describedby={item.ID}>{item.Title}</span>
          </TooltipHost>
        </>
      ),
    },
    {
      key: "Column4",
      name: "N/A",
      fieldName: "NotApplicable",
      minWidth: 50,
      maxWidth: 50,
      onRender: (item, Index) => (
        <Checkbox
          styles={{ root: { marginTop: 3 } }}
          data-id={item.ID}
          //styles={apModalBoxCheckBoxStyles}
          disabled={!dpUpdate ? true : false}
          checked={item.NotApplicable}
          onChange={(ev) => {
            dpOnchangeItems(item.RefId, "NotApplicable", ev.target["checked"]);
          }}
        />
      ),
    },
    {
      key: "Column5",
      name: "N/A(M)",
      fieldName: "NotApplicableManager",
      minWidth: 50,
      maxWidth: 50,
      onRender: (item, Index) => (
        <Checkbox
          styles={{ root: { marginTop: 3 } }}
          data-id={item.ID}
          //styles={apModalBoxCheckBoxStyles}
          disabled={
            dpUpdate &&
            (apCurrentData.length > 0 && loggeduseremail != ""
              ? apCurrentData[0].ProjectOwnerEmail == loggeduseremail
              : null)
              ? false
              : true
          }
          checked={item.NotApplicableManager}
          onChange={(ev) => {
            dpOnchangeItems(
              item.RefId,
              "NotApplicableManager",
              ev.target["checked"]
            );
          }}
        />
      ),
    },
    {
      key: "Column6",
      name: "Hours",
      fieldName: "PlannedHours",
      minWidth: 50,
      maxWidth: 50,
      onRender: (item, Index) => (
        <TextField
          styles={{
            root: {
              selectors: {
                ".ms-TextField-fieldGroup": {
                  borderRadius: 4,
                  border: "1px solid",
                  height: 28,
                  input: {
                    borderRadius: 4,
                  },
                },
              },
            },
          }}
          data-id={item.ID}
          disabled={!dpUpdate || item.Source == "DP" ? true : false}
          value={item.PlannedHours}
          onChange={(e: any) => {
            dpOnchangeItems(item.RefId, "PlannedHours", e.target.value);
          }}
        />
      ),
    },
    {
      key: "Column7",
      name: "Start date",
      fieldName: "StartDate",
      minWidth: 120,
      maxWidth: 120,
      onRender: (item, Index) => (
        <DatePicker
          data-id={item.ID}
          // label="Start Date"
          placeholder="Select a date..."
          ariaLabel="Select a date"
          formatDate={dateFormater}
          // minDate={new Date(apCurrentData[0].StartDate)}
          // maxDate={new Date(apCurrentData[0].PlannedEndDate)}
          styles={{
            textField: {
              transform: "translateY(3px)",
              selectors: {
                ".ms-TextField-fieldGroup": {
                  borderColor: item.DateError ? "#d0342c" : "#000",
                  borderRadius: 4,
                  border: "1px solid",
                  height: 27,
                  input: {
                    borderRadius: 4,
                  },
                },
                ".ms-TextField-field": {
                  color: item.DateError ? "#d0342c" : "#000",
                },
                ".ms-DatePicker-event--without-label": {
                  color: item.DateError ? "#d0342c" : "#000",
                  paddingTop: 3,
                },
              },
            },
            readOnlyTextField: {
              lineHeight: 25,
            },
          }}
          value={item.StartDate ? new Date(item.StartDate) : new Date()}
          disabled={!dpUpdate ? true : false}
          onSelectDate={(value: any) => {
            dpOnchangeItems(item.RefId, "StartDate", value);
            let refIndex = dpData.findIndex((obj) => obj.RefId == item.RefId);
            if (
              moment(value).format("MM/DD/YYYY") <=
              moment(item.EndDate).format("MM/DD/YYYY")
            ) {
              dpData[refIndex].DateError = false;
              dpDateErrorFunction();
            } else {
              dpData[refIndex].DateError = true;
              dpDateErrorFunction();
            }
          }}
        />
      ),
    },
    {
      key: "Column8",
      name: "End date",
      fieldName: "EndDate",
      minWidth: 120,
      maxWidth: 120,
      onRender: (item, Index) => (
        <DatePicker
          data-id={item.ID}
          // label="Start Date"
          placeholder="Select a date..."
          ariaLabel="Select a date"
          formatDate={dateFormater}
          // minDate={new Date(apCurrentData[0].StartDate)}
          // maxDate={new Date(apCurrentData[0].PlannedEndDate)}
          value={item.EndDate ? new Date(item.EndDate) : new Date()}
          disabled={!dpUpdate ? true : false}
          styles={{
            textField: {
              transform: "translateY(3px)",
              selectors: {
                ".ms-TextField-fieldGroup": {
                  borderColor: item.DateError ? "#d0342c" : "#000",
                  borderRadius: 4,
                  border: "1px solid",
                  height: 27,
                  input: {
                    borderRadius: 4,
                  },
                },
                ".ms-TextField-field": {
                  color: item.DateError ? "#d0342c" : "#000",
                },
                ".ms-DatePicker-event--without-label": {
                  color: item.DateError ? "#d0342c" : "#000",
                  paddingTop: 3,
                },
              },
            },
            readOnlyTextField: {
              lineHeight: 25,
            },
          }}
          //styles={apModalBoxDatePickerStyles}
          onSelectDate={(value: any) => {
            dpOnchangeItems(item.RefId, "EndDate", value);
            let refIndex = dpData.findIndex((obj) => obj.RefId == item.RefId);
            if (
              moment(item.StartDate).format("MM/DD/YYYY") <=
              moment(value).format("MM/DD/YYYY")
            ) {
              dpData[refIndex].DateError = false;
              dpDateErrorFunction();
            } else {
              dpData[refIndex].DateError = true;
              dpDateErrorFunction();
            }
          }}
        />
      ),
    },
    {
      key: "Column9",
      name: "Status",
      fieldName: "Status",
      minWidth: 100,
      maxWidth: 100,
      onRender: (item) => (
        <div /*style={{ marginTop: "0.2rem" }}*/>
          {item.Status == "Completed" ? (
            <div className={dpstatusStyleClass.completed}>{item.Status}</div>
          ) : item.Status == "Scheduled" ? (
            <div className={dpstatusStyleClass.scheduled}>{item.Status}</div>
          ) : item.Status == "On schedule" ? (
            <div className={dpstatusStyleClass.onSchedule}>{item.Status}</div>
          ) : item.Status == "Behind schedule" ? (
            <div className={dpstatusStyleClass.behindScheduled}>
              {item.Status}
            </div>
          ) : (
            ""
          )}
        </div>
      ),
    },
    {
      key: "Column10",
      name: "Developer",
      fieldName: "Developer",
      minWidth: 250,
      maxWidth: 250,
      onRender: (item, Index) => (
        <NormalPeoplePicker
          styles={{
            root: {
              selectors: {
                ".ms-BasePicker-text": {
                  height: 28,
                  padding: 1,
                  border: "1px solid #000",
                  borderRadius: 4,
                },
              },
            },
          }}
          data-id={item.ID}
          //className={apModalBoxPP}
          onResolveSuggestions={GetUserDetails}
          itemLimit={1}
          disabled={!dpUpdate ? true : false}
          // defaultSelectedItems={item.defaultDeveloper}
          //selectedItems={item.defaultDeveloper}
          selectedItems={
            item.DeveloperId != null
              ? [
                  peopleList.filter((people) => {
                    return people.ID == item.DeveloperId;
                  })[0],
                ]
              : []
          }
          onChange={(selectedUser) => {
            selectedUser.length != 0
              ? dpOnchangeItems(
                  item.RefId,
                  "DeveloperId",
                  selectedUser[0]["ID"]
                )
              : dpOnchangeItems(item.RefId, "DeveloperId", null);
          }}
        />
      ),
    },
    {
      key: "Column11",
      name: "TL",
      fieldName: "TLink",
      minWidth: 30,
      maxWidth: 30,
      onRender: (item) => (
        <>
          <a target="_blank" href={item.TLink}>
            <Icon
              iconName="NavigateExternalInline"
              className={dpiconStyleClass.link}
              style={{ color: "#038387" }}
            />
          </a>
        </>
      ),
    },
    {
      key: "Column12",
      name: "PBL",
      fieldName: "PBLink",
      minWidth: 40,
      maxWidth: 40,
      onRender: (item) => (
        <>
          <a target="_blank" href={item.PBLink}>
            <Icon
              iconName="NavigateExternalInline"
              className={dpiconStyleClass.link}
              style={{ color: "#038387" }}
            />
          </a>
        </>
      ),
    },
    {
      key: "Column13",
      name: "Action",
      fieldName: "Arrow",
      minWidth: 40,
      maxWidth: 40,
      onRender: (item, index) => (
        <div
          style={{
            display: "flex",
            flexDirection: "row",
            alignItems: "center",
            // justifyContent: "center",
            marginTop: 6,
          }}
        >
          {index != 0 ? (
            <Icon
              iconName="ChevronUpMed"
              className={dpiconStyleClass.link}
              style={{
                cursor:
                  dpUpdate && dpData.length == dpDisplayData.length
                    ? "pointer"
                    : "default",
                color: "#038387",
                marginRight: 12,
              }}
              onClick={(_) => {
                dpUpdate && dpData.length == dpDisplayData.length
                  ? dpArrows(index, "Up")
                  : null;
              }}
            />
          ) : (
            // <img
            //   src={`${upArrowImg}`}
            //   alt="up"
            //   style={{
            //     cursor:
            //       dpUpdate && dpData.length == dpDisplayData.length
            //         ? "pointer"
            //         : "default",
            //     color: "#038387",
            //     marginRight: 5,
            //     transform: "scale(1.2)",
            //   }}
            //   onClick={(_) => {
            //     dpUpdate && dpData.length == dpDisplayData.length
            //       ? dpArrows(index, "Up")
            //       : null;
            //   }}
            // />
            ""
          )}
          {index != dpData.length - 1 ? (
            <Icon
              iconName="ChevronDownMed"
              className={dpiconStyleClass.link}
              style={{
                cursor:
                  dpUpdate && dpData.length == dpDisplayData.length
                    ? "pointer"
                    : "default",
                color: "#038387",
              }}
              onClick={(_) => {
                dpUpdate && dpData.length == dpDisplayData.length
                  ? dpArrows(index, "Down")
                  : null;
              }}
            />
          ) : (
            // <img
            //   src={`${downArrowImg}`}
            //   alt="down"
            //   style={{
            //     cursor:
            //       dpUpdate && dpData.length == dpDisplayData.length
            //         ? "pointer"
            //         : "default",
            //     color: "#038387",
            //     width: 20,
            //   }}
            //   onClick={(_) => {
            //     dpUpdate && dpData.length == dpDisplayData.length
            //       ? dpArrows(index, "Down")
            //       : null;
            //   }}
            // />
            ""
          )}
        </div>
      ),
    },
  ];
  const dpDrpDwnOptns = {
    source: [{ key: "All", text: "All" }],
    // activity: [{ key: "All", text: "All" }],
    status: [{ key: "All", text: "All" }],
    developer: [{ key: "All", text: "All" }],
  };
  const dpFilterKeys = {
    source: "All",
    // activity: "All",
    status: "All",
    developer: "All",
  };
  let dpErrorStatus = {
    Deliverable: "",
    Source: "",
  };
  const dpAddItems = {
    RefId: 0,
    ID: 0,
    AnnualPlanID: Ap_AnnualPlanId,
    Source: "CIM",
    ProductId: "",
    Title: "",
    NotApplicable: false,
    NotApplicableManager: false,
    StartDate: null,
    EndDate: null,
    Status: "Scheduled",
    DeveloperId: null,
    ManagerId: null,
    PlannedHours: 0,
    Week: Dp_WeekNumber,
    Year: Dp_Year,
    TLink: "https://www.w3schools.com",
    PBLink: "https://www.w3schools.com",
    DateError: false,
    BA: null,
    ActualHours: 0,
    Onchange: true,
  };
  const Source = [
    { key: "CIM", text: "CIM" },
    { key: "OM", text: "OM" },
  ];
  // Design
  const dpProjectInfo = mergeStyles({
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    margin: "0 15px",
  });
  const dplabelStyles = mergeStyleSets({
    titleLabel: [
      {
        color: "#676767",
        fontSize: "14px",
        marginRight: "10px",
        fontWeight: "400",
      },
    ],
    labelValue: [
      {
        color: "#0882A5",
        fontSize: "14px",
        marginRight: "10px",
      },
    ],
    inputLabels: [
      {
        color: "#323130",
        fontSize: "13px",
      },
    ],
    ErrorLabel: [
      {
        marginTop: "25px",
        marginLeft: "10px",
        fontWeight: "500",
        color: "#D0342C",
        fontSize: "13px",
      },
    ],
    NORLabel: [
      {
        marginTop: "25px",
        marginLeft: "10px",
        fontWeight: "500",
        color: "#323130",
        fontSize: "13px",
      },
    ],
  });

  const dpdropdownStyles: Partial<IDropdownStyles> = {
    root: { width: 186, marginRight: 15 },
    title: {
      backgroundColor: "#F5F5F7",
      fontSize: 12,
      border: "1px solid #E8E8EA",
    },
    dropdownItemsWrapper: { backgroundColor: "#F5F5F7", fontSize: 12 },
    dropdownItemSelected: { backgroundColor: "#DCDCDC", fontSize: 12 },
    caretDown: { fontSize: 14, color: "#000" },
  };
  const dpiconStyle = mergeStyles({
    fontSize: 17,
    height: 14,
    width: 17,
    cursor: "pointer",
  });
  const dpiconStyleClass = mergeStyleSets({
    link: [{ color: "blue", margin: "0 0" }, dpiconStyle],
    delete: [{ color: "red", margin: "0 7px" }, dpiconStyle],
    edit: [{ color: "blue", margin: "0 7px" }, dpiconStyle],
    refresh: [
      {
        color: "white",
        fontSize: "18px",
        height: 22,
        width: 22,
        cursor: "pointer",
        backgroundColor: "#038387",
        padding: 5,
        marginTop: 28,
        borderRadius: 2,
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
        ":hover": {
          backgroundColor: "#025d60",
        },
      },
    ],
    pblink: [
      {
        color: "white",
        fontSize: "18px",
        height: 22,
        width: 22,
        cursor: "pointer",
        backgroundColor: "#038387",
        padding: 5,
        marginLeft: 10,
        borderRadius: 2,
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
      },
    ],
  });
  const dpBigiconStyle = mergeStyles({
    fontSize: 25,
    height: 20,
    width: 25,
    cursor: "pointer",
    marginRight: 10,
    marginTop: 2,
  });
  const dpBigiconStyleClass = mergeStyleSets({
    ChevronLeftMed: [
      {
        cursor: "pointer",
        color: "#2392b2",
        fontSize: 24,
        marginTop: "3px",
        marginRight: 12,
      },
    ],
  });
  const dpprofileName = mergeStyles({
    // color: "#E2A00F",
    fontWeight: "bold",
    marginRight: "10px",
  });

  // detailslist
  const gridStyles: Partial<IDetailsListStyles> = {
    root: {
      // overflowX: "scroll",
      selectors: {
        "& [role=grid]": {
          display: "flex",
          flexDirection: "column",
          // alignItems: "start",
          // height: "60vh",
          // ".ms-DetailsRow-cell": {
          //   height: 42,
          //   padding: "6px 12px",
          // },
          ".ms-DetailsRow-fields": {
            alignItems: "center",
            // height: 42,
          },
        },
      },
    },
    headerWrapper: {
      flex: "0 0 auto",
    },
    contentWrapper: {
      flex: "1 1 auto",
      overflowY: "auto",
      overflowX: "hidden",
    },
  };

  const dpstatusStyle = mergeStyles({
    textAlign: "center",
    paddingTop: 2,
    borderRadius: "25px",
  });
  const dpstatusStyleClass = mergeStyleSets({
    completed: [
      {
        fontWeight: "600",
        padding: "3px",
        color: "#438700",
        backgroundColor: "#D9FFB3",
      },
      dpstatusStyle,
    ],
    scheduled: [
      {
        fontWeight: "600",
        padding: "3px",
        color: "#9f6700",
        backgroundColor: "#FFDB99",
      },
      dpstatusStyle,
    ],
    onSchedule: [
      {
        fontWeight: "600",
        padding: "3px",
        color: "#B3B300 ",
        backgroundColor: "#FFFFB3",
      },
      dpstatusStyle,
    ],
    behindScheduled: [
      {
        fontWeight: "600",
        padding: "3px",
        color: "#FF0000",
        backgroundColor: "#FFB3B3",
      },
      dpstatusStyle,
    ],
  });
  const dpbuttonStyle = mergeStyles({
    textAlign: "center",
    borderRadius: "2px",
  });
  const dpbuttonStyleClass = mergeStyleSets({
    buttonPrimary: [
      {
        color: "White",
        backgroundColor: "#FAA332",
        borderRadius: "3px",
        border: "none",
        marginRight: "10px",
        selectors: {
          ":hover": {
            backgroundColor: "#FAA332",
            opacity: 0.9,
            borderRadius: "3px",
            border: "none",
            marginRight: "10px",
          },
        },
      },
      dpbuttonStyle,
    ],
    buttonSecondary: [
      {
        color: "White",
        backgroundColor: "#038387",
        borderRadius: "3px",
        border: "none",
        margin: "0 5px",
        selectors: {
          ":hover": {
            backgroundColor: "#038387",
            opacity: 0.9,
          },
        },
      },
      dpbuttonStyle,
    ],
  });
  const dpTxtBoxStyles: Partial<ITextFieldStyles> = {
    root: { width: 400, marginLeft: 25, marginTop: 15 },
    field: { backgroundColor: "whitesmoke", fontSize: 12 },
  };
  //-----------------------------------Heading Styles ---------------------------------//
  const dpTxtHeadingBoxStyles: Partial<ITextFieldStyles> = {
    root: {
      width: 450,
      margin: 0,
      backgroundColor: "#2392B2",
      textAlign: "center",
      height: 40,
      padding: 10,
      fontSize: 18,
      fontWeight: 600,
      color: "White",
    },
  };

  // useState
  const [dpUpdate, setdpUpdate] = useState(false);
  const [dpReRender, setdpReRender] = useState(false);
  const [apCurrentData, setapCurrentData] = useState([]);
  const [dpData, setdpData] = useState([]);
  const [dpMasterData, setdpMasterData] = useState([]);
  const [dpDisplayData, setdpDisplayData] = useState([]);
  const [dpDeliverable, setdpDeliverable] = useState(dpAddItems);
  const [dpDropDownOptions, setdpDropDownOptions] = useState(dpDrpDwnOptns);
  const [dpFilterOptions, setdpFilterOptions] = useState(dpFilterKeys);
  const [peopleList, setPeopleList] = useState(allPeoples);
  const [dpModalBoxVisibility, setdpModalBoxVisibility] = useState(false);
  const [dpPopup, setdpPopup] = useState("");
  const [dpShowMessage, setdpShowMessage] = useState(dpErrorStatus);
  const [dpLoader, setdpLoader] = useState(true);
  const [dpErrorDate, setdpErrorDate] = useState(false);
  const [dpButtonLoader, setdpButtonLoader] = useState(false);
  const [dpAutoSave, setdpAutoSave] = useState(false);
  const [thisweekPBData, setthisweekPBData] = useState([]);

  const dateFormater = (date: Date): string => {
    return !date ? "" : moment(date).format("DD/MM/YYYY");
  };

  // Get All site users
  const getAllUsers = () => {
    sharepointWeb.siteUsers().then((_allUsers) => {
      _allUsers.forEach((user) => {
        allPeoples.push({
          key: 1,
          imageUrl:
            `/_layouts/15/userphoto.aspx?size=S&accountname=` + `${user.Email}`,
          text: user.Title,
          ID: user.Id,
          secondaryText: user.Email,
          isValid: true,
        });
      });
      setPeopleList(allPeoples);
    });
  };
  const GetUserDetails = (filterText) => {
    var result = peopleList.filter(
      (value, index, self) => index === self.findIndex((t) => t.ID === value.ID)
    );
    return result.filter((item) =>
      doesTextStartWith(item.text as string, filterText)
    );
  };
  const doesTextStartWith = (text: string, filterText: string) => {
    return text.toLowerCase().indexOf(filterText.toLowerCase()) === 0;
  };

  // Getting data from Delivery Plan List and Dropdown Options
  const getthisweekPBData = () => {
    sharepointWeb.lists
      .getByTitle("ProductionBoard")
      .items.filter(
        "Week eq '" +
          Dp_WeekNumber +
          "' and AnnualPlanID eq '" +
          Ap_AnnualPlanId +
          "' "
      )
      .get()
      .then(async (items) => {
        setthisweekPBData([...items]);
      })
      .catch(dpErrorFunction);
  };

  const getCurrentAPData = () => {
    let _apCurrentData = [];

    sharepointWeb.lists
      .getByTitle(ListNameURL)
      .items //.getById(Ap_AnnualPlanId)
      .select(
        "*",
        "ProjectOwner/Title",
        "ProjectOwner/Id",
        "ProjectOwner/EMail",
        "ProjectLead/Title",
        "ProjectLead/Id",
        "ProjectLead/EMail",
        "Master_x0020_Project/Title",
        "Master_x0020_Project/Id"
      )
      .expand("ProjectOwner", "ProjectLead", "Master_x0020_Project")
      .filter("ID eq '" + Ap_AnnualPlanId + "' ")
      .get()
      .then(async (items) => {
        console.log(items);
        items.forEach((item) => {
          _apCurrentData.push({
            ID: item.ID,
            Title: item.Title,
            TypeofProject: item.ProjectType,
            ProductId: item.Master_x0020_ProjectId,
            ProductName: item.Master_x0020_Project
              ? item.Master_x0020_Project.Title
              : "",
            DeveloperId: item.ProjectLeadId ? item.ProjectLeadId[0] : null,
            ProjectOwnerId: item.ProjectOwnerId,
            DeveloperEmail: item.ProjectLead ? item.ProjectLead[0].EMail : null,
            ProjectOwnerEmail: item.ProjectOwner
              ? item.ProjectOwner.EMail
              : null,
            DeveloperName: item.ProjectLead ? item.ProjectLead[0].Title : null,
            ProjectOwnerName: item.ProjectOwner
              ? item.ProjectOwner.Title
              : null,
            StartDate: item.StartDate,
            PlannedEndDate: item.PlannedEndDate,
            AllocatedHours: item.AllocatedHours ? item.AllocatedHours : 0,
            BA: item.BusinessArea,
            Onchange: false,
          });
        });

        setapCurrentData([..._apCurrentData]);
        getDpData(_apCurrentData[0]);
      })
      .catch(dpErrorFunction);
  };
  const getDpData = (_apData) => {
    sharepointWeb.lists
      .getByTitle("Delivery Plan")
      .items.select(
        "*,Developer/Title,Developer/Id,Developer/EMail,Manager/Title,Manager/EMail"
      )
      .expand("Developer, Manager")
      .filter("AnnualPlanID eq '" + Ap_AnnualPlanId + "' ")
      .orderBy("OrderId", true)
      .top(5000)
      .get()
      .then((items) => {
        //console.log(items);
        items.forEach((item: any, index: number) => {
          _dpAllitems.push({
            RefId: index + 1,
            ID: item.ID,
            Source: item.Source,
            Title: item.Title,
            NotApplicable: item.NotApplicable,
            NotApplicableManager: item.NotApplicableManager,
            StartDate: item.StartDate
              ? moment(item.StartDate).format("MM/DD/yyyy")
              : null,
            EndDate: item.EndDate
              ? moment(item.EndDate).format("MM/DD/yyyy")
              : null,
            Status: item.Status,
            DeveloperId: item.DeveloperId,
            ManagerId: item.ManagerId ? item.ManagerId : null,
            PlannedHours: item.PlannedHours ? item.PlannedHours : 0,
            Week: item.Week,
            Year: item.Year,
            //Hours: item.PlannedHours,
            TLink: "https://www.w3schools.com",
            PBLink: "https://www.w3schools.com",
            DateError: false,
            BA: item.BA,
            ActualHours: item.ActualHours ? item.ActualHours : 0,
            Onchange: false,
          });
        });

        if (_dpAllitems.length == 0) {
          getDpTemplateData(_apData);
        } else {
          setdpLoader(false);
          setdpDropDownOptions(dpDrpDwnOptns);
          setdpMasterData([..._dpAllitems]);
          setdpData([..._dpAllitems]);
          setdpDisplayData([..._dpAllitems]);
          reloadFilterOptions(_dpAllitems);
        }
      })
      .catch(dpErrorFunction);
  };
  const getDpTemplateData = (_apData) => {
    sharepointWeb.lists
      .getByTitle("Delivery Plan Phase List")
      .items.filter("DeliverPlanTypeOfWork eq '" + _apData.TypeofProject + "' ")
      .top(5000)
      .get()
      .then((items) => {
        //console.log(items);
        items.forEach((item: any, index: number) => {
          _dpAllitems.push({
            RefId: index + 1,
            ID: 0,
            AnnualPlanID: Ap_AnnualPlanId,
            Source: "DP",
            ProductId: _apData.ProductId,
            Title: item.Title,
            NotApplicable: false,
            NotApplicableManager: false,
            StartDate: _apData.StartDate ? _apData.StartDate : null,
            EndDate: _apData.PlannedEndDate ? _apData.PlannedEndDate : null,
            Status: "Scheduled",
            DeveloperId: _apData.DeveloperId ? _apData.DeveloperId : null,
            ManagerId: _apData.ProjectOwnerId ? _apData.ProjectOwnerId : null,
            PlannedHours: item.Hours ? item.Hours : 0,
            Week: Dp_WeekNumber,
            Year: Dp_Year,
            TLink: "https://www.w3schools.com",
            PBLink: "https://www.w3schools.com",
            DateError: false,
            BA: _apData.BA,
            ActualHours: 0,
            Onchange: false,
          });
        });
        if (_dpAllitems.length > 0) {
          setdpDropDownOptions(dpDrpDwnOptns);
          setdpMasterData([..._dpAllitems]);
          setdpData([..._dpAllitems]);
          setdpDisplayData([..._dpAllitems]);
          reloadFilterOptions(_dpAllitems);
        } else {
          setdpLoader(false);
        }
      })
      .catch(dpErrorFunction);
  };
  const reloadFilterOptions = (data) => {
    let tempArrReload = data;

    tempArrReload.forEach((dp) => {
      if (
        dpDrpDwnOptns.source.findIndex((source) => {
          return source.key == dp.Source;
        }) == -1
      ) {
        dpDrpDwnOptns.source.push({
          key: dp.Source,
          text: dp.Source,
        });
      }
      if (
        dpDrpDwnOptns.status.findIndex((status) => {
          return status.key == dp.Status;
        }) == -1
      ) {
        dpDrpDwnOptns.status.push({
          key: dp.Status,
          text: dp.Status,
        });
      }
      let devName =
        peopleList.length > 0 && dp.DeveloperId != null
          ? peopleList.filter((people) => {
              return people.ID == dp.DeveloperId;
            })[0].text
          : "";
      if (
        dpDrpDwnOptns.developer.findIndex((developer) => {
          return developer.key == dp.DeveloperId;
        }) == -1 &&
        devName != ""
      ) {
        dpDrpDwnOptns.developer.push({
          key: dp.DeveloperId,
          text: devName,
        });
      }
    });
    setdpLoader(false);
    setdpDropDownOptions(dpDrpDwnOptns);
  };

  // Button Click Function
  const cancelDPData = () => {
    reloadFilterOptions(dpMasterData);
    setdpDisplayData([...dpMasterData]);
    setdpData([...dpMasterData]);
    setdpUpdate(false);
  };
  const saveDPData = () => {
    setdpLoader(true);
    let Hours = sumOfHours();
    let successCount = 0;
    dpData.forEach((dp, Index: number) => {
      let requestdata = {
        AnnualPlanIDId: Ap_AnnualPlanId,
        Source: dp.Source ? dp.Source : null,
        Title: dp.Title ? dp.Title : null,
        StartDate: dp.StartDate
          ? moment(dp.StartDate).format("MM/DD/yyyy")
          : null,
        EndDate: dp.EndDate ? moment(dp.EndDate).format("MM/DD/yyyy") : null,
        DeveloperId: dp.DeveloperId ? dp.DeveloperId : null,
        ManagerId: dp.ManagerId ? dp.ManagerId : null,
        NotApplicable: dp.NotApplicable ? dp.NotApplicable : null,
        NotApplicableManager: dp.NotApplicableManager
          ? dp.NotApplicableManager
          : null,
        Status: dp.Status ? dp.Status : null,
        PlannedHours: dp.PlannedHours ? dp.PlannedHours : 0,
        ProductId: dp.ProductId ? dp.ProductId : null,
        Week: Dp_WeekNumber,
        Year: Dp_Year,
        OrderId: Index,
        BA: dp.BA ? dp.BA : null,
        AnnualPlanIDNumber: Ap_AnnualPlanId,
        ActualHours: dp.ActualHours ? dp.ActualHours : 0,
      };
      if (dp.ID != 0) {
        sharepointWeb.lists
          .getByTitle("Delivery Plan")
          .items.getById(dp.ID)
          .update(requestdata)
          .then((e) => {
            successCount++;

            let updatePB = thisweekPBData.filter((pb) => {
              return pb.DeliveryPlanID == dp.ID;
            });

            if (updatePB.length > 0 && dp.Onchange == true) {
              sharepointWeb.lists
                .getByTitle("ProductionBoard")
                .items.getById(updatePB[0].ID)
                .update({
                  StartDate: dp.StartDate
                    ? moment(dp.StartDate).format("MM/DD/yyyy")
                    : null,
                  EndDate: dp.EndDate
                    ? moment(dp.EndDate).format("MM/DD/yyyy")
                    : null,
                  DeveloperId: dp.DeveloperId ? dp.DeveloperId : null,
                  NotApplicable: dp.NotApplicable,
                  NotApplicableManager: dp.NotApplicableManager,
                })
                .then((e) => {})
                .catch(dpErrorFunction);
            }

            if (dpData.length == successCount) {
              sharepointWeb.lists
                .getByTitle(ListNameURL)
                .items.getById(Ap_AnnualPlanId)
                .update({ AllocatedHours: Hours })
                .then((e) => {
                  apCurrentData[0].AllocatedHours = Hours;
                })
                .catch(dpErrorFunction);

              setdpMasterData([...dpData]);
              setdpUpdate(false);
              setdpPopup("Success");
              setTimeout(() => {
                setdpPopup("Close");
              }, 2000);
              setdpLoader(false);
            }
          })
          .catch(dpErrorFunction);
      } else if (dp.ID == 0) {
        sharepointWeb.lists
          .getByTitle("Delivery Plan")
          .items.add(requestdata)
          .then((e) => {
            successCount++;
            dpData[Index].ID = e.data.ID;

            if (thisweekPBData.length > 0 && dp.Source != "DP") {
              sharepointWeb.lists
                .getByTitle("ProductionBoard")
                .items.add({
                  BA: dp.BA ? dp.BA : null,
                  StartDate: dp.StartDate
                    ? moment(dp.StartDate).format("MM/DD/yyyy")
                    : null,
                  EndDate: dp.EndDate
                    ? moment(dp.EndDate).format("MM/DD/yyyy")
                    : null,
                  Source: dp.Source ? dp.Source : null,
                  AnnualPlanIDId: dp.AnnualPlanID ? dp.AnnualPlanID : null,
                  ProductId: dp.ProductId ? dp.ProductId : null,
                  Title: dp.Title ? dp.Title : null,
                  PlannedHours: dp.PlannedHours ? dp.PlannedHours : null,
                  Monday: 0,
                  Tuesday: 0,
                  Wednesday: 0,
                  Thursday: 0,
                  Friday: 0,
                  ActualHours: 0,
                  DeveloperId: dp.DeveloperId ? dp.DeveloperId : null,
                  Week: Dp_WeekNumber,
                  Year: Dp_Year,
                  NotApplicable: dp.NotApplicable,
                  NotApplicableManager: dp.NotApplicableManager,
                  DeliveryPlanID: dp.ID,
                  DPActualHours: 0,
                  Status: "Pending",
                  AnnualPlanIDNumber: dp.AnnualPlanID,
                })
                .then((e) => {})
                .catch(dpErrorFunction);
            }

            if (dpData.length == successCount) {
              sharepointWeb.lists
                .getByTitle(ListNameURL)
                .items.getById(Ap_AnnualPlanId)
                .update({ AllocatedHours: Hours })
                .then((e) => {
                  apCurrentData[0].AllocatedHours = Hours;
                })
                .catch(dpErrorFunction);

              setdpData([...dpData]);
              setdpMasterData([...dpData]);
              setdpUpdate(false);
              setdpPopup("Success");
              setTimeout(() => {
                setdpPopup("Close");
              }, 2000);
              setdpLoader(false);
            }
          })
          .catch(dpErrorFunction);
      }
    });
  };
  // Onchange and Filters
  const dpListFilter = (key, option) => {
    let tempArr = [...dpData];
    let tempDpFilterKeys = { ...dpFilterOptions };
    tempDpFilterKeys[`${key}`] = option;

    if (tempDpFilterKeys.source != "All") {
      tempArr = tempArr.filter((arr) => {
        return arr.Source == tempDpFilterKeys.source;
      });
    }
    // if (tempDpFilterKeys.activity != "All") {
    //   tempArr = tempArr.filter((arr) => {
    //     return arr.Title == tempDpFilterKeys.activity;
    //   });
    // }
    if (tempDpFilterKeys.status != "All") {
      tempArr = tempArr.filter((arr) => {
        return arr.Status == tempDpFilterKeys.status;
      });
    }
    if (tempDpFilterKeys.developer != "All") {
      tempArr = tempArr.filter((arr) => {
        return arr.DeveloperId == tempDpFilterKeys.developer;
      });
    }

    // setdpData([...tempArr]);
    setdpDisplayData([...tempArr]);
    setdpFilterOptions({ ...tempDpFilterKeys });
  };
  const dpAddOnchange = (key, value) => {
    let tempArronchange = dpDeliverable;
    if (key == "deliverable") dpDeliverable.Title = value;
    else if (key == "source") dpDeliverable.Source = value;

    dpDeliverable.ManagerId = apCurrentData[0].ProjectOwnerId;
    dpDeliverable.DeveloperId = apCurrentData[0].DeveloperId;
    dpDeliverable.ProductId = apCurrentData[0].ProductId;
    dpDeliverable.BA = apCurrentData[0].BA;
    dpDeliverable.EndDate = apCurrentData[0].PlannedEndDate
      ? apCurrentData[0].PlannedEndDate
      : new Date();
    dpDeliverable.StartDate = apCurrentData[0].StartDate
      ? apCurrentData[0].StartDate
      : new Date();
    dpDeliverable.RefId = dpData.length + 1;
    setdpDeliverable(tempArronchange);
  };
  const dpArrows = (Index, Type) => {
    if (Type == "Up") {
      var srcUp = dpData[Index];
      var desUp = dpData[Index - 1];

      dpData[Index] = desUp;
      dpData[Index - 1] = srcUp;

      setdpDisplayData([...dpData]);
      setdpData([...dpData]);
    }
    if (Type == "Down") {
      var srcDown = dpData[Index];
      var desDown = dpData[Index + 1];

      dpData[Index] = desDown;
      dpData[Index + 1] = srcDown;

      setdpDisplayData([...dpData]);
      setdpData([...dpData]);
    }
  };
  const dpOnchangeItems = (RefId, key, value) => {
    let Index = dpData.findIndex((obj) => obj.RefId == RefId);
    let disIndex = dpDisplayData.findIndex((obj) => obj.RefId == RefId);
    let dpBeforeData = dpData[Index];
    let dpOnchangeData = [
      {
        RefId: dpBeforeData.RefId,
        ID: dpBeforeData.ID,
        AnnualPlanID: dpBeforeData.AnnualPlanID,
        Source: dpBeforeData.Source,
        ProductId: dpBeforeData.ProductId,
        Title: dpBeforeData.Title,
        NotApplicable:
          key == "NotApplicable" ? value : dpBeforeData.NotApplicable,
        NotApplicableManager:
          key == "NotApplicableManager"
            ? value
            : dpBeforeData.NotApplicableManager,
        StartDate: key == "StartDate" ? value : dpBeforeData.StartDate,
        EndDate: key == "EndDate" ? value : dpBeforeData.EndDate,
        Status: dpBeforeData.Status,
        DeveloperId: key == "DeveloperId" ? value : dpBeforeData.DeveloperId,
        ManagerId: dpBeforeData.ManagerId,
        PlannedHours: key == "PlannedHours" ? value : dpBeforeData.PlannedHours,
        Week: dpBeforeData.Week,
        Year: dpBeforeData.Year,
        TLink: dpBeforeData.TLink,
        PBLink: dpBeforeData.PBLink,
        DateError: dpBeforeData.DateError,
        BA: dpBeforeData.BA,
        ActualHours: dpBeforeData.ActualHours,
        Onchange: true,
      },
    ];

    dpData[Index] = dpOnchangeData[0];
    dpDisplayData[disIndex] = dpOnchangeData[0];
    reloadFilterOptions(dpData);
    setdpData([...dpData]);
  };

  // Header Content
  const sumOfHours = () => {
    var sum: number = 0;
    if (dpData.length > 0) {
      dpData.forEach((x) => {
        sum += parseInt(x.PlannedHours ? x.PlannedHours : 0);
      });
      return sum;
    }
  };
  const sumOfActualHours = () => {
    var sum: number = 0;
    if (dpData.length > 0) {
      dpData.forEach((x) => {
        sum += parseInt(x.ActualHours ? x.ActualHours : 0);
      });
      return sum;
    }
  };
  const overallStatus = () => {
    let status = "Scheduled";
    const resultCompleted = dpData.filter((dp) => dp.Status == "Completed");
    if (resultCompleted.length == dpData.length) {
      status = "Completed";
    } else {
      status = "Scheduled";
    }
    return (
      <div style={{ width: "100px", marginTop: "4px" }}>
        {status == "Completed" ? (
          <div className={dpstatusStyleClass.completed}>{status}</div>
        ) : status == "Scheduled" ? (
          <div className={dpstatusStyleClass.scheduled}>{status}</div>
        ) : (
          ""
        )}
      </div>
    );
  };

  // Validation and Success
  const dpValidationFunction = () => {
    let isError = false;

    let errorStatus = {
      Deliverable: "",
      Source: "",
    };

    if (!dpDeliverable.Title) {
      isError = true;
      errorStatus.Deliverable = "Please enter a value for deliverable";
    }
    if (!dpDeliverable.Source) {
      isError = true;
      errorStatus.Source = "Please Select a value for Source";
    }

    if (!isError) {
      setdpButtonLoader(true);
      setdpDisplayData(dpDisplayData.concat(dpDeliverable));
      setdpData(dpData.concat(dpDeliverable));
      setdpModalBoxVisibility(false);
      reloadFilterOptions(dpDisplayData.concat(dpDeliverable));
      setdpUpdate(true);
      console.log(dpData.concat(dpDeliverable));
    } else {
      setdpShowMessage(errorStatus);
    }
  };
  const dpErrorFunction = (error) => {
    setdpLoader(false);
    console.log(error);
    if (dpData.length > 0) {
      setdpPopup("Error");
      setTimeout(() => {
        setdpPopup("Close");
      }, 2000);
    }
  };
  const SuccessPopup = () => (
    <MessageBar messageBarType={MessageBarType.success} isMultiline={false}>
      Delivery plan has been Successfully Submitted !!!
    </MessageBar>
  );
  const ErrorPopup = () => (
    <MessageBar messageBarType={MessageBarType.error} isMultiline={false}>
      Something when Error, please contact system admin.
    </MessageBar>
  );
  const dpDateErrorFunction = () => {
    let resultDateError = dpData.filter((dp) => dp.DateError == true);

    if (resultDateError.length > 0) {
      setdpErrorDate(true);
    } else {
      setdpErrorDate(false);
    }
  };

  window.onbeforeunload = function (e) {
    debugger;
    if (dpAutoSave) {
      let dialogText =
        "You have unsaved changes, are you sure you want to leave?";
      e.returnValue = dialogText;
      return dialogText;
    }
  };

  const alertDialog = () => {
    if (confirm("You have unsaved changes, are you sure you want to leave?")) {
      props.handleclick("AnnualPlan", null, "DP");
      //console.log("Thing was saved to the database.");
    } else {
      //console.log("Thing was not saved to the database.");
    }
  };

  useEffect(() => {
    if (dpAutoSave && !dpErrorDate && dpUpdate) {
      setTimeout(() => {
        document.getElementById("btnSave").click();
      }, 300000);
    }
  }, [dpAutoSave]);

  useEffect(() => {
    getCurrentAPData();
    getAllUsers();
    getthisweekPBData();
  }, [dpReRender]);

  // return function are dispaly in page
  return (
    <div style={{ padding: "5px 15px" }}>
      {dpLoader ? (
        // <Spinner
        //   label="Please wait..."
        //   size={SpinnerSize.large}
        //   style={{
        //     width: "100vw",
        //     height: "100vh",
        //     position: "fixed",
        //     top: 0,
        //     left: 0,
        //     backgroundColor: "#c8c6c49c",
        //     zIndex: 10000,
        //   }}
        // />
        <CustomLoader />
      ) : null}
      <div className={styles.dpHeaderSection} style={{ paddingBottom: "0 " }}>
        <div
          style={{
            position: "sticky",
            top: 0,
            backgroundColor: "#fff",
            zIndex: 1,
            marginBottom: 41,
          }}
        >
          <>
            {dpPopup == "Success"
              ? SuccessPopup()
              : dpPopup == "Error"
              ? ErrorPopup()
              : ""}
          </>
          <div
            style={{
              display: "flex",
              alignItems: "flex-start",
              justifyContent: "space-between",
              marginBottom: 20,
              color: "#2392b2",
            }}
          >
            {/* Header Start */}
            <div className={styles.dpTitle}>
              <Icon
                aria-label="ChevronLeftMed"
                iconName="ChevronLeftMed"
                className={dpBigiconStyleClass.ChevronLeftMed}
                onClick={() => {
                  dpAutoSave
                    ? confirm(
                        "You have unsaved changes, are you sure you want to leave?"
                      )
                      ? props.handleclick("AnnualPlan", null, "DP")
                      : null
                    : props.handleclick("AnnualPlan", null, "DP");
                }}
              />
              {/* <img
              src={`${BackIconImg}`}
              alt="back"
              style={{
                width: 40,
                cursor: "pointer",
                marginRight: 10,
                marginTop: 5,
              }}
              onClick={() => {
                dpAutoSave ? alertDialog() : props.handleclick("AnnualPlan");
              }}
            /> */}
              <Label style={{ color: "#000", fontSize: 24, padding: 0 }}>
                Delivery plan
              </Label>
            </div>
            <div>
              {
                <div
                  title={
                    apCurrentData.length > 0
                      ? apCurrentData[0].ProjectOwnerName
                      : ""
                  }
                  className={styles.userDetails}
                >
                  <Label className={dpprofileName} style={{ display: "flex" }}>
                    <div style={{ color: "#E2A00F", width: 80 }}>Manager: </div>
                    {apCurrentData.length > 0
                      ? apCurrentData[0].ProjectOwnerName
                      : ""}
                  </Label>
                  <Persona
                    size={PersonaSize.size24}
                    presence={PersonaPresence.none}
                    // text={`${
                    //   apCurrentData.length > 0 ? apCurrentData[0].ProjectOwnerName : ""
                    // }`}
                    imageUrl={
                      "/_layouts/15/userphoto.aspx?size=S&username=" +
                      `${
                        apCurrentData.length > 0
                          ? apCurrentData[0].ProjectOwnerEmail
                          : ""
                      }`
                    }
                  />
                </div>
              }
              {
                <div
                  title={
                    apCurrentData.length > 0
                      ? apCurrentData[0].DeveloperName
                      : ""
                  }
                  className={styles.userDetails}
                >
                  <Label className={dpprofileName} style={{ display: "flex" }}>
                    <div style={{ color: "#E2A00F", width: 80 }}>
                      Developer:{" "}
                    </div>
                    {apCurrentData.length > 0
                      ? apCurrentData[0].DeveloperName
                      : ""}
                  </Label>
                  <Persona
                    size={PersonaSize.size24}
                    presence={PersonaPresence.none}
                    // text={`${
                    //   apCurrentData.length > 0 ? apCurrentData[0].DeveloperName : ""
                    // }`}
                    imageUrl={
                      "/_layouts/15/userphoto.aspx?size=S&username=" +
                      `${
                        apCurrentData.length > 0
                          ? apCurrentData[0].DeveloperEmail
                          : ""
                      }`
                    }
                  />
                </div>
              }
            </div>
          </div>

          <div
            style={{
              display: "flex",
              alignItems: "center",
              justifyContent: "space-between",
            }}
          >
            <div className={styles.Section1}>
              {apCurrentData.length > 0 &&
              apCurrentData[0].TypeofProject != null ? (
                <PrimaryButton
                  text="Add deliverable"
                  className={dpbuttonStyleClass.buttonPrimary}
                  onClick={(_) => {
                    setdpDeliverable(dpAddItems);
                    setdpShowMessage(dpErrorStatus);
                    setdpModalBoxVisibility(true);
                    setdpButtonLoader(false);
                  }}
                />
              ) : (
                <PrimaryButton
                  text="Add deliverable"
                  disabled={true}
                  //className={dpbuttonStyleClass.buttonPrimary}
                  onClick={(_) => {
                    // setdpDeliverable(dpAddItems);
                    // setdpShowMessage(dpErrorStatus);
                    // setdpModalBoxVisibility(true);
                    // setdpButtonLoader(false);
                  }}
                />
              )}
              <div className={dpProjectInfo}>
                <Label className={dplabelStyles.titleLabel}>
                  Project or task :
                </Label>
                <Label
                  className={dplabelStyles.labelValue}
                  style={{ maxWidth: 250 }}
                >
                  {apCurrentData.length > 0 ? apCurrentData[0].Title : ""}
                </Label>
              </div>
              <div className={dpProjectInfo}>
                <Label className={dplabelStyles.titleLabel}>Product :</Label>
                <Label
                  className={dplabelStyles.labelValue}
                  style={{ maxWidth: 250 }}
                >
                  {apCurrentData.length > 0 ? apCurrentData[0].ProductName : ""}
                </Label>
              </div>
              <div className={dpProjectInfo}>
                <Label className={dplabelStyles.titleLabel}>AH/hrs :</Label>
                <Label className={dplabelStyles.labelValue}>
                  {sumOfActualHours()} / {sumOfHours()}
                </Label>
                {/* <Label className={dplabelStyles.labelValue}> / </Label>
              <Label className={dplabelStyles.labelValue}>{sumOfHours()}</Label> */}
              </div>
              <div className={dpProjectInfo}>
                <Label className={dplabelStyles.titleLabel}>Status :</Label>
                {overallStatus()}
              </div>
              <div className={dpProjectInfo}>
                <Label className={dplabelStyles.titleLabel}>TOD :</Label>
                <Label className={dplabelStyles.labelValue}>
                  {apCurrentData.length > 0
                    ? apCurrentData[0].TypeofProject
                    : null}
                </Label>
              </div>
            </div>
            {dpData.length > 0 ? (
              <div>
                <div
                  style={{
                    display: "flex",
                    alignItems: "center",
                    justifyContent: "center",
                  }}
                >
                  {dpUpdate ? (
                    <PrimaryButton
                      iconProps={cancelIcon}
                      text="Cancel"
                      className={dpbuttonStyleClass.buttonPrimary}
                      onClick={(_) => {
                        setdpAutoSave(false);
                        setdpErrorDate(false);
                        cancelDPData();
                      }}
                    />
                  ) : (
                    <PrimaryButton
                      iconProps={editIcon}
                      text="Edit"
                      className={dpbuttonStyleClass.buttonPrimary}
                      onClick={(_) => {
                        setdpAutoSave(true);
                        setdpUpdate(true);
                      }}
                    />
                  )}

                  {dpErrorDate || !dpUpdate ? (
                    <PrimaryButton
                      iconProps={saveIcon}
                      text="Save"
                      disabled={true}
                      //className={dpbuttonStyleClass.buttonSecondary}
                      onClick={(_) => {
                        setdpAutoSave(false);
                        saveDPData();
                      }}
                    />
                  ) : (
                    <PrimaryButton
                      iconProps={saveIcon}
                      id="btnSave"
                      text="Save"
                      className={dpbuttonStyleClass.buttonSecondary}
                      onClick={(_) => {
                        setdpAutoSave(false);
                        saveDPData();
                      }}
                    />
                  )}
                  <Icon
                    iconName="Link12"
                    className={dpiconStyleClass.pblink}
                    onClick={() => {
                      dpAutoSave
                        ? confirm(
                            "You have unsaved changes, are you sure you want to leave?"
                          )
                          ? props.handleclick(
                              "ProductionBoard",
                              Ap_AnnualPlanId,
                              "DP"
                            )
                          : null
                        : props.handleclick(
                            "ProductionBoard",
                            Ap_AnnualPlanId,
                            "DP"
                          );
                    }}
                  />
                </div>
              </div>
            ) : null}
          </div>

          <div
            style={{
              display: "flex",
              alignItems: "center",
              justifyContent: "space-between",
              marginBottom: "15px",
              paddingBottom: "10px",
            }}
          >
            <div className={styles.ddSection}>
              {/* Section Start */}
              <div>
                <Label className={dplabelStyles.inputLabels}>Source</Label>
                <Dropdown
                  selectedKey={dpFilterOptions.source}
                  placeholder="Select an option"
                  options={dpDropDownOptions.source}
                  styles={dpdropdownStyles}
                  onChange={(e, option: any) => {
                    dpListFilter("source", option["key"]);
                  }}
                />
              </div>
              {/* <div>
          <Label >Activity</Label>
          <Dropdown
            defaultSelectedKey={"All"}
            placeholder="Select an option"
            options={dpDropDownOptions.activity}
            styles={dpdropdownStyles}
            // onChange={(e, option: any) => {
            //   dpListFilter("activity", option["key"]);
            // }}
          />
        </div> */}

              <div>
                <Label className={dplabelStyles.inputLabels}>Status</Label>
                <Dropdown
                  selectedKey={dpFilterOptions.status}
                  placeholder="Select an option"
                  options={dpDropDownOptions.status}
                  styles={dpdropdownStyles}
                  onChange={(e, option: any) => {
                    dpListFilter("status", option["key"]);
                  }}
                />
              </div>
              <div>
                <Label className={dplabelStyles.inputLabels}>Developer</Label>
                <Dropdown
                  selectedKey={dpFilterOptions.developer}
                  placeholder="Select an option"
                  options={dpDropDownOptions.developer}
                  styles={dpdropdownStyles}
                  onChange={(e, option: any) => {
                    dpListFilter("developer", option["key"]);
                  }}
                />
              </div>
              <div>
                <Icon
                  iconName="Refresh"
                  className={dpiconStyleClass.refresh}
                  onClick={() => {
                    if (dpAutoSave) {
                      if (
                        confirm(
                          "You have unsaved changes, are you sure you want to leave?"
                        )
                      ) {
                        setdpData([...dpMasterData]);
                        setdpDisplayData([...dpMasterData]);
                        setdpFilterOptions({ ...dpFilterKeys });
                        setdpUpdate(false);
                      }
                    } else {
                      setdpData([...dpMasterData]);
                      setdpDisplayData([...dpMasterData]);
                      setdpFilterOptions({ ...dpFilterKeys });
                      setdpUpdate(false);
                    }
                  }}
                />
              </div>

              {dpErrorDate ? (
                <div>
                  <Label className={dplabelStyles.ErrorLabel}>
                    Given end date should not be earlier than the start date
                  </Label>
                </div>
              ) : null}
              {/* Section Start */}
            </div>

            <div>
              <Label className={dplabelStyles.NORLabel}>
                Number of records:{" "}
                <b style={{ color: "#038387" }}>{dpData.length}</b>
              </Label>
            </div>
          </div>
        </div>
        {/* Header- End */}
      </div>
      <div style={{ marginTop: -40 }}>
        {/* dont remove */}
        <input
          id="forFocus"
          type="text"
          style={{
            width: 0,
            height: 0,
            border: "none",
            position: "absolute",
            top: 0,
            left: 0,
            padding: 0,
          }}
        />
      </div>
      <div
        className={styles.scrollTop}
        onClick={() => {
          document.querySelector("#forFocus")["focus"]();
          // window.location.hash = "#idForFocus";
        }}
      >
        <Icon
          iconName="Up"
          className={dpiconStyleClass.link}
          style={{ color: "#fff" }}
        />
      </div>
      <div>
        {
          <DetailsList
            items={dpDisplayData}
            columns={_dpColumns}
            setKey="set"
            layoutMode={DetailsListLayoutMode.justified}
            selectionMode={SelectionMode.none}
            styles={gridStyles}
            // styles={{
            //   root: {
            //     selectors: {
            //       ".ms-DetailsRow-cell": {
            //         height: 42,
            //         padding: "6px 12px",
            //       },
            //     },
            //   },
            // }}
          />
        }
      </div>
      {dpData.length == 0 ? (
        <div
          style={{
            display: "flex",
            justifyContent: "center",
            marginTop: "15px",
          }}
        >
          <Label>No data Found !!!</Label>
        </div>
      ) : null}

      <div>
        <Modal
          // titleAriaId={titleId}
          isOpen={dpModalBoxVisibility}
          // onDismiss={hideModal}
          isBlocking={false}
          // containerClassName={contentStyles.container}
          // dragOptions={isDraggable ? dragOptions : undefined}
        >
          <div>
            {" "}
            <Label styles={dpTxtHeadingBoxStyles}>Add deliverable</Label>
          </div>
          <TextField
            styles={dpTxtBoxStyles}
            required={true}
            errorMessage={dpShowMessage.Deliverable}
            label="Deliverable"
            onChange={(e, value: string) => {
              dpAddOnchange("deliverable", value);
            }}
          />
          <div>
            <ChoiceGroup
              styles={dpTxtBoxStyles}
              defaultSelectedKey="CIM"
              options={Source}
              label="Source"
              onChange={(e, option: any) => {
                dpAddOnchange("source", option["key"]);
              }}
            />
          </div>
          <div></div>
          <div className={styles.apModalBoxButtonSection}>
            <button
              className={styles.apModalBoxSubmitBtn}
              onClick={(_) => {
                setdpAutoSave(true);
                dpValidationFunction();
                document.querySelector("#forFocusBottom")["focus"]();
              }}
            >
              {dpButtonLoader ? (
                <Spinner />
              ) : (
                // <CustomLoader />
                <span>
                  <Icon
                    iconName="Save"
                    style={{ marginTop: 4, marginRight: 12 }}
                  />
                  {"Submit"}
                </span>
              )}
            </button>
            <button
              className={styles.apModalBoxBackBtn}
              onClick={(_) => {
                setdpShowMessage(dpErrorStatus);
                setdpDeliverable(dpAddItems);
                setdpModalBoxVisibility(false);
              }}
            >
              <span>
                {" "}
                <Icon
                  iconName="ChromeBack"
                  style={{ marginTop: 4, marginRight: 12 }}
                />
                Back
              </span>
            </button>
          </div>
        </Modal>
      </div>
      <div>
        {/* dont remove */}
        <input
          id="forFocusBottom"
          type="text"
          style={{
            width: 0,
            height: 0,
            border: "none",
            padding: 20,
          }}
        />
      </div>
    </div>
  );
}

export default DeliveryPlan;
