import * as React from "react";
import { useState, useEffect } from "react";
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
  PrimaryButton,
  TextField,
  ITextFieldStyles,
  MessageBar,
  MessageBarType,
  Spinner,
  SpinnerSize,
  ILabelStyles,
  Toggle,
  Modal,
  NormalPeoplePicker,
  TooltipHost,
  TooltipOverflowMode,
} from "@fluentui/react";
import * as Excel from "exceljs/dist/exceljs.min.js";
import * as FileSaver from "file-saver";
import { FontIcon, IIconProps } from "@fluentui/react/lib/Icon";
import "../ExternalRef/styleSheets/ProdStyles.css";
import { mergeStyles, mergeStyleSets } from "@fluentui/react/lib/Styling";
import styles from "./InnovationHubIntranet.module.scss";
import CustomLoader from "./CustomLoader";
import Pagination from "office-ui-fabric-react-pagination";
import { IDetailsListStyles } from "office-ui-fabric-react";

function ProductionBoard(props: any) {
  const sharepointWeb = Web(props.URL);
  const ListNameURL = props.WeblistURL;
  let Ap_AnnualPlanId = props.AnnualPlanId;
  let navType = props.pageType;

  let Pb_Year = moment().year();
  let Pb_NextWeekYear = moment().add(1, "week").year();
  let Pb_LastWeekYear = moment().subtract(1, "week").year();

  let Pb_WeekNumber = moment().isoWeek();
  let Pb_NextWeekNumber = moment().add(1, "week").isoWeek();
  let Pb_LastWeekNumber = moment().subtract(1, "week").isoWeek();

  let thisWeekMonday = moment().day(1).format("YYYY-MM-DD");
  let thisWeekTuesday = moment().day(2).format("YYYY-MM-DD");
  let thisWeekWednesday = moment().day(3).format("YYYY-MM-DD");
  let thisWeekThursday = moment().day(4).format("YYYY-MM-DD");
  let thisWeekFriday = moment().day(5).format("YYYY-MM-DD");

  let loggeduseremail = props.spcontext.pageContext.user.email;
  let currentpage = 1;
  let totalPageItems = 10;
  const allPeoples = [];

  // Initialization function
  const drAllitems = {
    Request: null,
    Requestto: null,
    Emailcc: null,
    Project: null,
    Documenttype: null,
    Link: null,
    Comments: null,
    Confidential: false,
    Product: null,
    AnnualPlanID: null,
    DeliveryPlanID: null,
    ProductionBoardID: null,
  };

  const pbFilterKeys = {
    BA: "All",
    Source: "All",
    Product: "All",
    Project: "All",
    Showonly: "Mine",
    Week: "This Week",
  };
  let pbErrorStatus = {
    Request: "",
    Requestto: "",
    Documenttype: "",
    Link: "",
  };
  const pbDrpDwnOptns = {
    BA: [{ key: "All", text: "All" }],
    Source: [{ key: "All", text: "All" }],
    Product: [{ key: "All", text: "All" }],
    Project: [{ key: "All", text: "All" }],
    Showonly: [
      //{ key: "All", text: "All" },
      { key: "Mine", text: "Mine" },
      { key: "All BA", text: "All BA" },
    ],
    Week: [
      //{ key: "All", text: "All" },
      { key: "This Week", text: "This Week" },
      { key: "Last Week", text: "Last Week" },
      { key: "Next Week", text: "Next Week" },
    ],
  };
  const pbModalBoxDrpDwnOptns = {
    Request: [],
    Documenttype: [],
  };
  const BAacronymsCollection = [
    {
      Name: "PD Curriculum",
      ShortName: "PDC",
    },
    {
      Name: "PD Professional Learning",
      ShortName: "PDPL",
    },
    {
      Name: "PD School Improvements",
      ShortName: "PDSI",
    },
    {
      Name: "SS Business",
      ShortName: "SSB",
    },
    {
      Name: "SS Publishing",
      ShortName: "SSP",
    },
    {
      Name: "SS Content Creation",
      ShortName: "SSCC",
    },
    {
      Name: "SS Marketing",
      ShortName: "SSM",
    },
    {
      Name: "SS Technology",
      ShortName: "SST",
    },
    {
      Name: "SS Research and Evaluation",
      ShortName: "SSRE",
    },
    {
      Name: "SD School Partnerships",
      ShortName: "SSPSP",
    },
  ];

  //Detail list Columns
  const _dpColumns = [
    {
      key: "Column1",
      name: "BA",
      fieldName: "BA",
      minWidth: 40,
      maxWidth: 40,
      onRender: (item) =>
        BAacronymsCollection.filter((BAacronym) => {
          return BAacronym.Name == item.BA;
        })[0].ShortName,
    },
    {
      key: "Column2",
      name: "Start Date",
      fieldName: "StartDate",
      minWidth: 70,
      maxWidth: 70,
      onRender: (item) => moment(item.StartDate).format("DD/MM/YYYY"),
    },
    {
      key: "Column3",
      name: "End Date",
      fieldName: "EndDate",
      minWidth: 70,
      maxWidth: 70,
      onRender: (item) => moment(item.EndDate).format("DD/MM/YYYY"),
    },
    {
      key: "Column4",
      name: "Source",
      fieldName: "Source",
      minWidth: 50,
      maxWidth: 50,
    },
    {
      key: "Column5",
      name: "Project or task",
      fieldName: "Project",
      minWidth: 160,
      maxWidth: 160,
      onRender: (item) => (
        <>
          <TooltipHost
            id={item.ID}
            content={item.Project}
            overflowMode={TooltipOverflowMode.Parent}
          >
            <span aria-describedby={item.ID}>{item.Project}</span>
          </TooltipHost>
        </>
      ),
    },
    {
      key: "Column6",
      name: "Product",
      fieldName: "Product",
      minWidth: 200,
      maxWidth: 200,
      onRender: (item) => (
        <>
          <TooltipHost
            id={item.ID}
            content={item.Product}
            overflowMode={TooltipOverflowMode.Parent}
          >
            <span aria-describedby={item.ID}>{item.Product}</span>
          </TooltipHost>
        </>
      ),
    },
    {
      key: "Column7",
      name: "Activity",
      fieldName: "Title",
      minWidth: 160,
      maxWidth: 160,
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
      key: "Column8",
      name: "PH",
      fieldName: "PlannedHours",
      minWidth: 30,
      maxWidth: 30,
    },
    {
      key: "Column9",
      name: "Mon",
      fieldName: "Monday",
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
          disabled={
            pbUpdate &&
            item.DeveloperEmail == loggeduseremail &&
            thisWeekMonday >= moment(item.StartDate).format("YYYY-MM-DD") &&
            thisWeekMonday <= moment(item.EndDate).format("YYYY-MM-DD")
              ? false
              : true
          }
          value={item.Monday}
          onChange={(e: any) => {
            pbOnchangeItems(item.RefId, "Monday", e.target.value);
          }}
        />
      ),
    },
    {
      key: "Column10",
      name: "Tue",
      fieldName: "Tuesday",
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
          disabled={
            pbUpdate &&
            item.DeveloperEmail == loggeduseremail &&
            thisWeekTuesday >= moment(item.StartDate).format("YYYY-MM-DD") &&
            thisWeekTuesday <= moment(item.EndDate).format("YYYY-MM-DD")
              ? false
              : true
          }
          value={item.Tuesday}
          onChange={(e: any) => {
            pbOnchangeItems(item.RefId, "Tuesday", e.target.value);
          }}
        />
      ),
    },
    {
      key: "Column11",
      name: "Wed",
      fieldName: "Wednesday",
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
          disabled={
            pbUpdate &&
            item.DeveloperEmail == loggeduseremail &&
            thisWeekWednesday >= moment(item.StartDate).format("YYYY-MM-DD") &&
            thisWeekWednesday <= moment(item.EndDate).format("YYYY-MM-DD")
              ? false
              : true
          }
          value={item.Wednesday}
          onChange={(e: any) => {
            pbOnchangeItems(item.RefId, "Wednesday", e.target.value);
          }}
        />
      ),
    },
    {
      key: "Column12",
      name: "Thu",
      fieldName: "Thursday",
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
          disabled={
            pbUpdate &&
            item.DeveloperEmail == loggeduseremail &&
            thisWeekThursday >= moment(item.StartDate).format("YYYY-MM-DD") &&
            thisWeekThursday <= moment(item.EndDate).format("YYYY-MM-DD")
              ? false
              : true
          }
          value={item.Thursday}
          onChange={(e: any) => {
            pbOnchangeItems(item.RefId, "Thursday", e.target.value);
          }}
        />
      ),
    },
    {
      key: "Column13",
      name: "Fri",
      fieldName: "Friday",
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
          disabled={
            pbUpdate &&
            item.DeveloperEmail == loggeduseremail &&
            thisWeekFriday >= moment(item.StartDate).format("YYYY-MM-DD") &&
            thisWeekFriday <= moment(item.EndDate).format("YYYY-MM-DD")
              ? false
              : true
          }
          value={item.Friday}
          onChange={(e: any) => {
            pbOnchangeItems(item.RefId, "Friday", e.target.value);
          }}
        />
      ),
    },
    {
      key: "Column14",
      name: "AH",
      fieldName: "ActualHours",
      minWidth: 50,
      maxWidth: 50,
    },
    {
      key: "Column15",
      name: "Action",
      fieldName: "",
      minWidth: 50,
      maxWidth: 50,
      onRender: (item, Index) =>
        pbWeek == Pb_WeekNumber &&
        item.DeveloperEmail == loggeduseremail &&
        item.ID != 0 ? (
          <Icon
            iconName="OpenEnrollment"
            //className={pbiconStyleClass.link}
            style={{
              color:
                item.Status == null
                  ? "#0882A5"
                  : item.Status == "Pending"
                  ? "#000000"
                  : item.Status == "Signed Off" ||
                    item.Status == "Published" ||
                    item.Status == "Completed"
                  ? "#40b200"
                  : item.Status == "Returned" || item.Status == "Cancelled"
                  ? "#ff3838"
                  : "#ffb302",
              marginTop: 6,
              marginLeft: 9,
              fontSize: 17,
              height: 14,
              width: 17,
              cursor: "pointer",
            }}
            onClick={(_) => {
              drAllitems.Project = item.Project;
              drAllitems.Product = item.Product;
              drAllitems.AnnualPlanID = item.AnnualPlanID;
              drAllitems.DeliveryPlanID = item.DeliveryPlanID;
              drAllitems.ProductionBoardID = item.ID;
              setpbButtonLoader(false);
              setpbShowMessage(pbErrorStatus);
              setpbDocumentReview(drAllitems);
              setpbModalBoxVisibility(true);
            }}
          />
        ) : (
          <Icon
            iconName="OpenEnrollment"
            //className={pbiconStyleClass.link}
            style={{
              color: "#ababab",
              marginTop: 6,
              marginLeft: 9,
              fontSize: 17,
              height: 14,
              width: 17,
              cursor: "default",
            }}
            onClick={(_) => {}}
          />
        ),
    },
  ];

  // Design
  const saveIcon: IIconProps = { iconName: "Save" };
  const editIcon: IIconProps = { iconName: "Edit" };
  const cancelIcon: IIconProps = { iconName: "Cancel" };
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
  const pbLabelStyles: Partial<ILabelStyles> = {
    root: {
      width: 150,
      marginRight: 10,
      fontSize: "13px",
      color: "#323130",
    },
  };
  const pbProjectInfo = mergeStyles({
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    margin: "0 10px",
  });
  const pblabelStyles = mergeStyleSets({
    titleLabel: [
      {
        color: "#676767",
        fontSize: "14px",
        marginRight: "10px",
        fontWeight: "400",
      },
    ],
    selectedLabel: [
      {
        color: "#0882A5",
        fontSize: "14px",
        marginRight: "10px",
        fontWeight: "600",
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
        color: "#323130",
        fontSize: "13px",
        marginLeft: "10px",
        fontWeight: "500",
      },
    ],
  });
  const pbBigiconStyleClass = mergeStyleSets({
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
  const pbbuttonStyle = mergeStyles({
    textAlign: "center",
    borderRadius: "2px",
  });
  const pbbuttonStyleClass = mergeStyleSets({
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
      pbbuttonStyle,
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
      pbbuttonStyle,
    ],
  });
  const pbiconStyle = mergeStyles({
    fontSize: 17,
    height: 14,
    width: 17,
    cursor: "pointer",
  });
  const pbiconStyleClass = mergeStyleSets({
    link: [{ color: "blue", margin: "0 0" }, pbiconStyle],
    delete: [{ color: "red", margin: "0 7px" }, pbiconStyle],
    edit: [{ color: "blue", margin: "0 7px" }, pbiconStyle],
    refresh: [
      {
        color: "white",
        fontSize: "18px",
        height: 22,
        width: 22,
        cursor: "pointer",
        backgroundColor: "#038387",
        padding: 5,
        marginTop: 27,
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
    export: [
      {
        color: "black",
        fontSize: "18px",
        height: 20,
        width: 20,
        cursor: "pointer",
        marginRight: 5,
      },
    ],
  });
  const pbDropdownStyles: Partial<IDropdownStyles> = {
    root: {
      width: 186,
      marginRight: "15px",
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
    // callout: { width: 300 },
  };
  const pbActiveDropdownStyles: Partial<IDropdownStyles> = {
    root: {
      width: 186,
      marginRight: "15px",
      backgroundColor: "#F5F5F7",
    },
    title: {
      backgroundColor: "#F5F5F7",
      fontSize: 12,
      color: "#038387",
      border: "2px solid #038387",
      borderRadius: "4px",
    },
    dropdownItem: {
      backgroundColor: "#F5F5F7",
      fontSize: 10,
    },
    dropdownItemSelected: {
      backgroundColor: "#F5F5F7",
      fontSize: 10,
    },
    caretDown: { fontSize: 14, color: "#000" },
  };

  const pbModalBoxDropdownStyles: Partial<IDropdownStyles> = {
    root: { width: "300px", margin: "10px 20px" },
    title: {
      fontSize: 12,
      borderRadius: "4px",
      border: "1px solid",
      height: "36px",
      padding: "3px 10px",
      color: "#000",
    },
    dropdownItemsWrapper: { fontSize: 12 },
    dropdownItemSelected: { fontSize: 12 },
    caretDown: {
      fontSize: 14,
      padding: "3px",
      color: "#000",
      fontWeight: "bold",
    },
    // label: { marginBottom: "3px" },
  };
  const pbModalBoxDrpDwnCalloutStyles: Partial<IDropdownStyles> = {
    root: { width: "300px", margin: "10px 20px" },
    title: {
      fontSize: 12,
      borderRadius: "4px",
      border: "1px solid",
      padding: "3px 10px",
      height: "36px",
      color: "#000",
    },
    dropdownItemsWrapper: { fontSize: 12 },
    dropdownItemSelected: { fontSize: 12 },
    caretDown: {
      fontSize: 14,
      paddingTop: "3px",
      color: "#000",
      fontWeight: "bold",
    },
    callout: { height: 200 },
    // label: { marginBottom: "3px" },
  };
  const pbTxtBoxStyles: Partial<ITextFieldStyles> = {
    root: {
      width: "300px",
      margin: "10px 20px",
    },
    field: {
      fontSize: 12,
      color: "#000",
      borderRadius: "4px",
      background: "#fff !important",
      // border: "1px solid !important",
    },
    fieldGroup: {
      border: "1px solid !important",
      height: "36px",
    },
  };
  const pbMultiTxtBoxStyles: Partial<ITextFieldStyles> = {
    root: {
      width: "640px",
      margin: "10px 20px",
      borderRadius: "4px",
    },
    field: { fontSize: 12, color: "#000" },
  };
  const pbModalBoxPP = mergeStyles({
    width: "300px",
    margin: "10px 20px",
  });

  // useState
  const [pbReRender, setpbReRender] = useState(false);
  const [pbChecked, setpbChecked] = useState(true);
  const [pbUpdate, setpbUpdate] = useState(false);
  const [pbDisplayData, setpbDisplayData] = useState([]);
  const [pbFilterData, setpbFilterData] = useState([]);
  const [pbData, setpbData] = useState([]);
  const [pbMasterData, setpbMasterData] = useState([]);
  const [pbDropDownOptions, setpbDropDownOptions] = useState(pbDrpDwnOptns);
  const [pbFilterOptions, setpbFilterOptions] = useState(pbFilterKeys);
  const [pbcurrentPage, setpbCurrentPage] = useState(currentpage);
  const [pbLoader, setpbLoader] = useState(true);
  const [pbModalBoxVisibility, setpbModalBoxVisibility] = useState(false);
  const [pbButtonLoader, setpbButtonLoader] = useState(false);
  const [pbModalBoxDropDownOptions, setpbModalBoxDropDownOptions] = useState(
    pbModalBoxDrpDwnOptns
  );
  const [peopleList, setPeopleList] = useState(allPeoples);
  const [pbDocumentReview, setpbDocumentReview] = useState(drAllitems);
  const [pbShowMessage, setpbShowMessage] = useState(pbErrorStatus);
  const [pbPopup, setpbPopup] = useState("");
  const [pbWeek, setpbWeek] = useState(Pb_WeekNumber);
  const [pbYear, setpbYear] = useState(Pb_Year);
  const [pbLastweek, setpbLastweek] = useState([]);
  const [pbNextweek, setpbNextweek] = useState([]);
  const [pbAutoSave, setpbAutoSave] = useState(false);
  // useEffect
  useEffect(() => {
    getAllUsers();
    getModalBoxOptions();
    getWeeksData("last", Pb_LastWeekNumber);
    getWeeksData("next", Pb_NextWeekNumber);

    Ap_AnnualPlanId ? getCurrentPbData() : getPbData();
  }, [pbReRender]);

  useEffect(() => {
    if (pbAutoSave && pbUpdate && pbWeek == Pb_WeekNumber) {
      setTimeout(() => {
        document.getElementById("btnSave").click();
      }, 300000);
    }
  }, [pbAutoSave]);

  window.onbeforeunload = function (e) {
    debugger;
    if (pbAutoSave) {
      let dialogText =
        "You have unsaved changes, are you sure you want to leave?";
      e.returnValue = dialogText;
      return dialogText;
    }
  };

  const alertDialogforBack = () => {
    if (confirm("You have unsaved changes, are you sure you want to leave?")) {
      navType == "AP"
        ? props.handleclick("AnnualPlan")
        : props.handleclick("DeliveryPlan", Ap_AnnualPlanId);
      //console.log("Thing was saved to the database.");
    } else {
      //console.log("Thing was not saved to the database.");
    }
  };

  // Functions
  const getWeeksData = (val, week) => {
    let _pbweek = [];
    if (Ap_AnnualPlanId) {
      sharepointWeb.lists
        .getByTitle("ProductionBoard")
        .items.select(
          "*,Developer/Title,Developer/Id,Developer/EMail,Product/Title,AnnualPlanID/Title"
        )
        .expand("Developer,Product,AnnualPlanID")
        .filter(
          "AnnualPlanID eq '" +
            Ap_AnnualPlanId +
            "' and Week eq '" +
            week +
            "' "
        )
        .top(5000)
        .get()
        .then((weekrecords) => {
          weekrecords.forEach((item, Index) => {
            _pbweek.push({
              RefId: Index + 1,
              ID: item.ID,
              BA: item.BA,
              StartDate: item.StartDate,
              EndDate: item.EndDate,
              Source: item.Source,
              Project: item.AnnualPlanID.Title,
              AnnualPlanID: item.AnnualPlanIDId,
              ProductId: item.ProductId,
              Product: item.Product ? item.Product.Title : "",
              Title: item.Title,
              PlannedHours: item.PlannedHours,
              Monday: item.Monday ? item.Monday : 0,
              Tuesday: item.Tuesday ? item.Tuesday : 0,
              Wednesday: item.Wednesday ? item.Wednesday : 0,
              Thursday: item.Thursday ? item.Thursday : 0,
              Friday: item.Friday ? item.Friday : 0,
              ActualHours: item.ActualHours,
              DeveloperId: item.DeveloperId,
              DeveloperEmail: item.Developer ? item.Developer.EMail : "",
              NotApplicable: item.NotApplicable,
              NotApplicableManager: item.NotApplicableManager,
              Week: item.Week,
              Year: item.Year,
              DeliveryPlanID: item.DeliveryPlanID,
              DPActualHours: item.DPActualHours,
              Status: item.Status,
            });
          });
          if (_pbweek.length == 0) {
            sharepointWeb.lists
              .getByTitle("Delivery Plan")
              .items.select(
                "*,Developer/Title,Developer/Id,Developer/EMail,Product/Title,AnnualPlanID/Title"
              )
              .expand("Developer,Product,AnnualPlanID")
              .filter("AnnualPlanID eq '" + Ap_AnnualPlanId + "' ")
              .top(5000)
              .get()
              .then((items) => {
                items.forEach((item, Index) => {
                  _pbweek.push({
                    RefId: Index + 1,
                    ID: 0,
                    BA: item.BA,
                    StartDate: item.StartDate,
                    EndDate: item.EndDate,
                    Source: item.Source,
                    Project: item.AnnualPlanID.Title,
                    AnnualPlanID: item.AnnualPlanIDId,
                    ProductId: item.ProductId,
                    Product: item.Product ? item.Product.Title : "",
                    Title: item.Title,
                    PlannedHours: item.PlannedHours,
                    Monday: 0,
                    Tuesday: 0,
                    Wednesday: 0,
                    Thursday: 0,
                    Friday: 0,
                    ActualHours: 0,
                    DeveloperId: item.DeveloperId,
                    DeveloperEmail: item.Developer ? item.Developer.EMail : "",
                    NotApplicable: item.NotApplicable,
                    NotApplicableManager: item.NotApplicableManager,
                    Week: Pb_WeekNumber,
                    Year: Pb_Year,
                    DeliveryPlanID: item.ID,
                    DPActualHours: item.ActualHours,
                    Status: null,
                  });
                });
                val == "last"
                  ? setpbLastweek([..._pbweek])
                  : setpbNextweek([..._pbweek]);
                //return [..._pbweek];
              })
              .catch(pbErrorFunction);
          } else {
            val == "last"
              ? setpbLastweek([..._pbweek])
              : setpbNextweek([..._pbweek]);
            //return [..._pbweek];
          }
        })
        .catch(pbErrorFunction);
    } else {
      sharepointWeb.lists
        .getByTitle("ProductionBoard")
        .items.select(
          "*,Developer/Title,Developer/Id,Developer/EMail,Product/Title,AnnualPlanID/Title"
        )
        .expand("Developer,Product,AnnualPlanID")
        .filter("Week eq '" + week + "' ")
        .top(5000)
        .get()
        .then((weekrecords) => {
          weekrecords.forEach((item, Index) => {
            _pbweek.push({
              RefId: Index + 1,
              ID: item.ID,
              BA: item.BA,
              StartDate: item.StartDate,
              EndDate: item.EndDate,
              Source: item.Source,
              Project: item.AnnualPlanID.Title,
              AnnualPlanID: item.AnnualPlanIDId,
              ProductId: item.ProductId,
              Product: item.Product ? item.Product.Title : "",
              Title: item.Title,
              PlannedHours: item.PlannedHours,
              Monday: item.Monday ? item.Monday : 0,
              Tuesday: item.Tuesday ? item.Tuesday : 0,
              Wednesday: item.Wednesday ? item.Wednesday : 0,
              Thursday: item.Thursday ? item.Thursday : 0,
              Friday: item.Friday ? item.Friday : 0,
              ActualHours: item.ActualHours,
              DeveloperId: item.DeveloperId,
              DeveloperEmail: item.Developer ? item.Developer.EMail : "",
              NotApplicable: item.NotApplicable,
              NotApplicableManager: item.NotApplicableManager,
              Week: item.Week,
              Year: item.Year,
              DeliveryPlanID: item.DeliveryPlanID,
              DPActualHours: item.DPActualHours,
              Status: item.Status,
            });
          });
          if (_pbweek.length == 0) {
            sharepointWeb.lists
              .getByTitle("Delivery Plan")
              .items.select(
                "*,Developer/Title,Developer/Id,Developer/EMail,Product/Title,AnnualPlanID/Title"
              )
              .expand("Developer,Product,AnnualPlanID")
              .filter("Developer/EMail eq '" + loggeduseremail + "' ")
              .top(5000)
              .get()
              .then(async (records) => {
                let count = 0;
                let getAnnualID = records.reduce(function (item, e1) {
                  var matches = item.filter(function (e2) {
                    return e1.AnnualPlanIDId === e2.AnnualPlanIDId;
                  });
                  if (matches.length == 0) {
                    item.push(e1);
                  }
                  return item;
                }, []);

                await getAnnualID.forEach(async (items) => {
                  await sharepointWeb.lists
                    .getByTitle("Delivery Plan")
                    .items.select(
                      "*,Developer/Title,Developer/Id,Developer/EMail,Product/Title,AnnualPlanID/Title"
                    )
                    .expand("Developer,Product,AnnualPlanID")
                    .filter("AnnualPlanID eq '" + items.AnnualPlanIDId + "' ")
                    .top(5000)
                    .get()
                    .then((items) => {
                      items.forEach((item) => {
                        _pbweek.push({
                          RefId: count++,
                          ID: 0,
                          BA: item.BA,
                          StartDate: item.StartDate,
                          EndDate: item.EndDate,
                          Source: item.Source,
                          Project: item.AnnualPlanID.Title,
                          AnnualPlanID: item.AnnualPlanIDId,
                          ProductId: item.ProductId,
                          Product: item.Product ? item.Product.Title : "",
                          Title: item.Title,
                          PlannedHours: item.PlannedHours,
                          Monday: 0,
                          Tuesday: 0,
                          Wednesday: 0,
                          Thursday: 0,
                          Friday: 0,
                          ActualHours: 0,
                          DeveloperId: item.DeveloperId,
                          DeveloperEmail: item.Developer
                            ? item.Developer.EMail
                            : "",
                          NotApplicable: item.NotApplicable,
                          NotApplicableManager: item.NotApplicableManager,
                          Week: Pb_WeekNumber,
                          Year: Pb_Year,
                          DeliveryPlanID: item.ID,
                          DPActualHours: item.ActualHours,
                          Status: null,
                        });
                      });
                      val == "last"
                        ? setpbLastweek([..._pbweek])
                        : setpbNextweek([..._pbweek]);
                      //return [..._pbweek];
                    });
                });
              })
              .catch(pbErrorFunction);
          } else {
            val == "last"
              ? setpbLastweek([..._pbweek])
              : setpbNextweek([..._pbweek]);
            //return [..._pbweek];
          }
        })
        .catch(pbErrorFunction);
    }
  };
  const getAllUsers = () => {
    sharepointWeb
      .siteUsers()
      .then((_allUsers) => {
        _allUsers.forEach((user) => {
          allPeoples.push({
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
        setPeopleList(allPeoples);
      })
      .catch(pbErrorFunction);
  };
  const getModalBoxOptions = () => {
    //Request Choices
    sharepointWeb.lists
      .getByTitle("ProductionBoard DR")
      .fields.getByInternalNameOrTitle("Request")()
      .then((response) => {
        response["Choices"].forEach((choice) => {
          if (choice != null) {
            if (
              pbModalBoxDrpDwnOptns.Request.findIndex((rpb) => {
                return rpb.key == choice;
              }) == -1
            ) {
              pbModalBoxDrpDwnOptns.Request.push({
                key: choice,
                text: choice,
              });
            }
          }
        });
      })
      .then()
      .catch(pbErrorFunction);

    //Documenttype Choices
    sharepointWeb.lists
      .getByTitle("ProductionBoard DR")
      .fields.getByInternalNameOrTitle("Documenttype")()
      .then((response) => {
        response["Choices"].forEach((choice) => {
          if (choice != null) {
            if (
              pbModalBoxDrpDwnOptns.Documenttype.findIndex((rdt) => {
                return rdt.key == choice;
              }) == -1
            ) {
              pbModalBoxDrpDwnOptns.Documenttype.push({
                key: choice,
                text: choice,
              });
            }
          }
        });
      })
      .then()
      .catch(pbErrorFunction);

    setpbModalBoxDropDownOptions(pbModalBoxDrpDwnOptns);
  };
  const generateExcel = () => {
    let arrExport = pbFilterData;
    const workbook = new Excel.Workbook();
    const worksheet = workbook.addWorksheet("My Sheet");
    worksheet.columns = [
      // { header: "ID", key: "id", width: 25 },
      { header: "Business area", key: "Businessarea", width: 25 },
      { header: "Start date", key: "Startdate", width: 25 },
      { header: "End date", key: "Enddate", width: 25 },
      { header: "Source", key: "Source", width: 25 },
      { header: "Project or task", key: "POT", width: 25 },
      { header: "Product", key: "Product", width: 60 },
      { header: "Activity", key: "Activity", width: 20 },
      { header: "PlannedHours", key: "PlannedHours", width: 40 },
      { header: "Monday", key: "Monday", width: 30 },
      { header: "Tuesday", key: "Tuesday", width: 30 },
      { header: "Wednesday", key: "Wednesday", width: 30 },
      { header: "Thursday", key: "Thursday", width: 30 },
      { header: "Friday", key: "Friday", width: 30 },
      { header: "Actual Total", key: "ActualTotal", width: 30 },
    ];
    arrExport.forEach((item) => {
      worksheet.addRow({
        // id: item.ID,
        Businessarea: item.BA ? item.BA : "",
        Startdate: item.StartDate ? item.StartDate : "",
        Enddate: item.EndDate ? item.EndDate : "",
        Source: item.Source ? item.Source : "",
        POT: item.Project ? item.Project : "",
        Product: item.Product ? item.Product : "",
        Activity: item.Title ? item.Title : "",
        PlannedHours: item.PlannedHours ? item.PlannedHours : "",
        Monday: item.Monday ? item.Monday : "",
        Tuesday: item.Tuesday ? item.Tuesday : "",
        Wednesday: item.Wednesday ? item.Wednesday : "",
        Thursday: item.Thursday ? item.Thursday : "",
        Friday: item.Friday ? item.Friday : "",
        ActualTotal: item.ActualHours ? item.ActualHours : "",
      });
    });
    [
      "A1",
      "B1",
      "C1",
      "D1",
      "E1",
      "F1",
      "G1",
      "H1",
      "I1",
      "J1",
      "K1",
      "L1",
      "M1",
      "N1",
    ].map((key) => {
      worksheet.getCell(key).fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "00e8d1" },
      };
    });
    [
      "A1",
      "B1",
      "C1",
      "D1",
      "E1",
      "F1",
      "G1",
      "H1",
      "I1",
      "J1",
      "K1",
      "L1",
      "M1",
      "N1",
    ].map((key) => {
      worksheet.getCell(key).color = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FFFFFF" },
      };
    });
    workbook.xlsx
      .writeBuffer()
      .then((buffer) =>
        FileSaver.saveAs(
          new Blob([buffer]),
          `ProductionBoard-${new Date().toLocaleString()}.xlsx`
        )
      )
      .catch((err) => console.log("Error writing excel export", err));
  };
  const getCurrentPbData = () => {
    sharepointWeb.lists
      .getByTitle("ProductionBoard")
      .items.select(
        "*,Developer/Title,Developer/Id,Developer/EMail,Product/Title,AnnualPlanID/Title"
      )
      .expand("Developer,Product,AnnualPlanID")
      .filter(
        "AnnualPlanID eq '" +
          Ap_AnnualPlanId +
          "' and Week eq '" +
          Pb_WeekNumber +
          "' "
      )
      .top(5000)
      .get()
      .then((items) => {
        let _pbAllitems = [];
        console.log(items);
        items.forEach((item, Index) => {
          _pbAllitems.push({
            RefId: Index + 1,
            ID: item.ID,
            BA: item.BA,
            StartDate: item.StartDate,
            EndDate: item.EndDate,
            Source: item.Source,
            Project: item.AnnualPlanID.Title,
            AnnualPlanID: item.AnnualPlanIDId,
            ProductId: item.ProductId,
            Product: item.Product ? item.Product.Title : "",
            Title: item.Title,
            PlannedHours: item.PlannedHours,
            Monday: item.Monday ? item.Monday : 0,
            Tuesday: item.Tuesday ? item.Tuesday : 0,
            Wednesday: item.Wednesday ? item.Wednesday : 0,
            Thursday: item.Thursday ? item.Thursday : 0,
            Friday: item.Friday ? item.Friday : 0,
            ActualHours: item.ActualHours,
            DeveloperId: item.DeveloperId,
            DeveloperEmail: item.Developer ? item.Developer.EMail : "",
            NotApplicable: item.NotApplicable,
            NotApplicableManager: item.NotApplicableManager,
            Week: item.Week,
            Year: item.Year,
            DeliveryPlanID: item.DeliveryPlanID,
            DPActualHours: item.DPActualHours ? item.DPActualHours : 0,
            Status: item.Status,
          });
        });

        if (_pbAllitems.length == 0) {
          getCurrentDpData();
        } else {
          setpbData([..._pbAllitems]);
          setpbMasterData([..._pbAllitems]);
          let pbFilter = ProductionBoardFilter(
            [..._pbAllitems],
            pbFilterOptions
          );
          reloadFilterOptions([...pbFilter]);
          setpbFilterData(pbFilter);
          paginate(1, pbFilter);
          setpbLoader(false);
        }
      })
      .catch(pbErrorFunction);
  };
  const getCurrentDpData = () => {
    sharepointWeb.lists
      .getByTitle("Delivery Plan")
      .items.select(
        "*,Developer/Title,Developer/Id,Developer/EMail,Product/Title,AnnualPlanID/Title"
      )
      .expand("Developer,Product,AnnualPlanID")
      .filter("AnnualPlanID eq '" + Ap_AnnualPlanId + "' ")
      .top(5000)
      .get()
      .then((items) => {
        let _pbAllitems = [];
        console.log(items);
        items.forEach((item, Index) => {
          _pbAllitems.push({
            RefId: Index + 1,
            ID: 0,
            BA: item.BA,
            StartDate: item.StartDate,
            EndDate: item.EndDate,
            Source: item.Source,
            Project: item.AnnualPlanID.Title,
            AnnualPlanID: item.AnnualPlanIDId,
            ProductId: item.ProductId,
            Product: item.Product ? item.Product.Title : "",
            Title: item.Title,
            PlannedHours: item.PlannedHours,
            Monday: 0,
            Tuesday: 0,
            Wednesday: 0,
            Thursday: 0,
            Friday: 0,
            ActualHours: 0,
            DeveloperId: item.DeveloperId,
            DeveloperEmail: item.Developer ? item.Developer.EMail : "",
            NotApplicable: item.NotApplicable,
            NotApplicableManager: item.NotApplicableManager,
            Week: Pb_WeekNumber,
            Year: Pb_Year,
            DeliveryPlanID: item.ID,
            DPActualHours: item.ActualHours ? item.ActualHours : 0,
            Status: null,
          });
        });
        setpbData([..._pbAllitems]);
        setpbMasterData([..._pbAllitems]);
        let pbFilter = ProductionBoardFilter([..._pbAllitems], pbFilterOptions);
        reloadFilterOptions([...pbFilter]);
        setpbFilterData(pbFilter);
        paginate(1, pbFilter);
        setpbLoader(false);
      })
      .catch(pbErrorFunction);
  };
  const getPbData = () => {
    sharepointWeb.lists
      .getByTitle("ProductionBoard")
      .items.select(
        "*,Developer/Title,Developer/Id,Developer/EMail,Product/Title,AnnualPlanID/Title"
      )
      .expand("Developer,Product,AnnualPlanID")
      .filter("Week eq '" + Pb_WeekNumber + "' ")
      .top(5000)
      .get()
      .then((items) => {
        let _pbAllitems = [];
        items.forEach((item, Index) => {
          _pbAllitems.push({
            RefId: Index + 1,
            ID: item.ID,
            BA: item.BA,
            StartDate: item.StartDate,
            EndDate: item.EndDate,
            Source: item.Source,
            Project: item.AnnualPlanID.Title,
            AnnualPlanID: item.AnnualPlanIDId,
            ProductId: item.ProductId,
            Product: item.Product ? item.Product.Title : "",
            Title: item.Title,
            PlannedHours: item.PlannedHours,
            Monday: item.Monday ? item.Monday : 0,
            Tuesday: item.Tuesday ? item.Tuesday : 0,
            Wednesday: item.Wednesday ? item.Wednesday : 0,
            Thursday: item.Thursday ? item.Thursday : 0,
            Friday: item.Friday ? item.Friday : 0,
            ActualHours: item.ActualHours,
            DeveloperId: item.DeveloperId,
            DeveloperEmail: item.Developer ? item.Developer.EMail : "",
            NotApplicable: item.NotApplicable,
            NotApplicableManager: item.NotApplicableManager,
            Week: item.Week,
            Year: item.Year,
            DeliveryPlanID: item.DeliveryPlanID,
            DPActualHours: item.DPActualHours ? item.DPActualHours : 0,
            Status: item.Status,
          });
        });
        if (_pbAllitems.length == 0) {
          getDpData();
        } else {
          setpbData([..._pbAllitems]);
          setpbMasterData([..._pbAllitems]);
          let pbFilter = ProductionBoardFilter(
            [..._pbAllitems],
            pbFilterOptions
          );
          reloadFilterOptions([...pbFilter]);
          setpbFilterData(pbFilter);
          paginate(1, pbFilter);
          setpbLoader(false);
        }
      })
      .catch(pbErrorFunction);
  };
  const getDpData = () => {
    sharepointWeb.lists
      .getByTitle("Delivery Plan")
      .items.select(
        "*,Developer/Title,Developer/Id,Developer/EMail,Product/Title,AnnualPlanID/Title"
      )
      .expand("Developer,Product,AnnualPlanID")
      .filter("Developer/EMail eq '" + loggeduseremail + "' ")
      .top(5000)
      .get()
      .then(async (records) => {
        let _pbAllitems = [];
        let count = 0;
        let getAnnualID = records.reduce(function (item, e1) {
          var matches = item.filter(function (e2) {
            return e1.AnnualPlanIDId === e2.AnnualPlanIDId;
          });
          if (matches.length == 0) {
            item.push(e1);
          }
          return item;
        }, []);

        await getAnnualID.forEach(async (items) => {
          await sharepointWeb.lists
            .getByTitle("Delivery Plan")
            .items.select(
              "*,Developer/Title,Developer/Id,Developer/EMail,Product/Title,AnnualPlanID/Title"
            )
            .expand("Developer,Product,AnnualPlanID")
            .filter("AnnualPlanID eq '" + items.AnnualPlanIDId + "' ")
            .top(5000)
            .get()
            .then((items) => {
              items.forEach((item) => {
                _pbAllitems.push({
                  RefId: count++,
                  ID: 0,
                  BA: item.BA,
                  StartDate: item.StartDate,
                  EndDate: item.EndDate,
                  Source: item.Source,
                  Project: item.AnnualPlanID.Title,
                  AnnualPlanID: item.AnnualPlanIDId,
                  ProductId: item.ProductId,
                  Product: item.Product ? item.Product.Title : "",
                  Title: item.Title,
                  PlannedHours: item.PlannedHours,
                  Monday: 0,
                  Tuesday: 0,
                  Wednesday: 0,
                  Thursday: 0,
                  Friday: 0,
                  ActualHours: 0,
                  DeveloperId: item.DeveloperId,
                  DeveloperEmail: item.Developer ? item.Developer.EMail : "",
                  NotApplicable: item.NotApplicable,
                  NotApplicableManager: item.NotApplicableManager,
                  Week: Pb_WeekNumber,
                  Year: Pb_Year,
                  DeliveryPlanID: item.ID,
                  DPActualHours: item.ActualHours ? item.ActualHours : 0,
                  Status: null,
                });
              });
              setpbData([..._pbAllitems]);
              setpbMasterData([..._pbAllitems]);
              let pbFilter = ProductionBoardFilter(
                [..._pbAllitems],
                pbFilterOptions
              );
              reloadFilterOptions([...pbFilter]);
              setpbFilterData(pbFilter);
              paginate(1, pbFilter);
              setpbLoader(false);
            });
        });
      })
      .catch(pbErrorFunction);
  };
  const savePBData = () => {
    setpbLoader(true);
    let successCount = 0;
    pbData.forEach((pb, Index: number) => {
      let requestdata = {
        BA: pb.BA,
        StartDate: pb.StartDate
          ? moment(pb.StartDate).format("MM/DD/yyyy")
          : null,
        EndDate: pb.EndDate ? moment(pb.EndDate).format("MM/DD/yyyy") : null,
        Source: pb.Source ? pb.Source : null,
        AnnualPlanIDId: pb.AnnualPlanID ? pb.AnnualPlanID : null,
        ProductId: pb.ProductId ? pb.ProductId : null,
        Title: pb.Title ? pb.Title : null,
        PlannedHours: pb.PlannedHours ? pb.PlannedHours : null,
        Monday: pb.Monday ? pb.Monday : 0,
        Tuesday: pb.Tuesday ? pb.Tuesday : 0,
        Wednesday: pb.Wednesday ? pb.Wednesday : 0,
        Thursday: pb.Thursday ? pb.Thursday : 0,
        Friday: pb.Friday ? pb.Friday : 0,
        ActualHours: pb.ActualHours ? pb.ActualHours : 0,
        DeveloperId: pb.DeveloperId ? pb.DeveloperId : null,
        Week: pb.Week,
        Year: pb.Year,
        NotApplicable: pb.NotApplicable,
        NotApplicableManager: pb.NotApplicableManager,
        DeliveryPlanID: pb.DeliveryPlanID,
        DPActualHours: pb.DPActualHours ? pb.DPActualHours : 0,
        Status: pb.Status ? pb.Status : null,
        AnnualPlanIDNumber: pb.AnnualPlanID,
      };
      let AH =
        parseInt(pb.DRActualHours ? pb.DRActualHours : 0) +
        parseInt(pb.ActualHours ? pb.ActualHours : 0);

      if (pb.ID != 0) {
        sharepointWeb.lists
          .getByTitle("ProductionBoard")
          .items.getById(pb.ID)
          .update(requestdata)
          .then(() => {
            sharepointWeb.lists
              .getByTitle("Delivery Plan")
              .items.getById(pb.DeliveryPlanID)
              .update({ ActualHours: AH })
              .then((e) => {})
              .catch(pbErrorFunction);

            successCount++;
            if (successCount == pbData.length) {
              setpbLoader(false);
              setpbMasterData([...pbData]);
              let pbFilter = ProductionBoardFilter([...pbData], pbFilterKeys);
              setpbFilterData(pbFilter);
              paginate(1, pbFilter);
              setpbFilterOptions({ ...pbFilterKeys });
              setpbUpdate(!pbUpdate);
              setpbPopup("Success");
              setTimeout(() => {
                setpbPopup("Close");
              }, 2000);
            }
          })
          .catch(pbErrorFunction);
      } else if (pb.ID == 0) {
        sharepointWeb.lists
          .getByTitle("ProductionBoard")
          .items.add(requestdata)
          .then((e) => {
            sharepointWeb.lists
              .getByTitle("Delivery Plan")
              .items.getById(pb.DeliveryPlanID)
              .update({ ActualHours: AH })
              .then((e) => {})
              .catch(pbErrorFunction);

            successCount++;
            pbData[Index].ID = e.data.ID;
            if (successCount == pbData.length) {
              setpbLoader(false);
              setpbData([...pbData]);
              setpbMasterData([...pbData]);
              let pbFilter = ProductionBoardFilter([...pbData], pbFilterKeys);
              setpbFilterData(pbFilter);
              paginate(1, pbFilter);
              setpbFilterOptions({ ...pbFilterKeys });
              setpbUpdate(!pbUpdate);
              setpbPopup("Success");
              setTimeout(() => {
                setpbPopup("Close");
              }, 2000);
            }
          })
          .catch(pbErrorFunction);
      }
    });
  };
  const savePBDRData = () => {
    let requestdata = {
      Title: pbDocumentReview.Link,
      Request: pbDocumentReview.Request ? pbDocumentReview.Request : null,
      RequesttoId: pbDocumentReview.Requestto
        ? pbDocumentReview.Requestto
        : null,
      EmailccId: pbDocumentReview.Emailcc
        ? { results: pbDocumentReview.Emailcc }
        : { results: [] },
      Project: pbDocumentReview.Project ? pbDocumentReview.Project : null,
      Documenttype: pbDocumentReview.Documenttype
        ? pbDocumentReview.Documenttype
        : null,
      Comments: pbDocumentReview.Comments ? pbDocumentReview.Comments : null,
      Confidential: pbDocumentReview.Confidential,
      Product: pbDocumentReview.Product ? pbDocumentReview.Product : null,
      AnnualPlanID: pbDocumentReview.AnnualPlanID
        ? pbDocumentReview.AnnualPlanID
        : null,
      DeliveryPlanID: pbDocumentReview.DeliveryPlanID
        ? pbDocumentReview.DeliveryPlanID
        : null,
      ProductionBoardID: pbDocumentReview.ProductionBoardID
        ? pbDocumentReview.ProductionBoardID
        : null,
    };
    sharepointWeb.lists
      .getByTitle("ProductionBoard DR")
      .items.add(requestdata)
      .then((e) => {
        if (pbDocumentReview.ProductionBoardID) {
          sharepointWeb.lists
            .getByTitle("ProductionBoard")
            .items.getById(pbDocumentReview.ProductionBoardID)
            .update({ Status: "Pending" })
            .then(() => {
              let Index = pbData.findIndex(
                (obj) => obj.ID == pbDocumentReview.ProductionBoardID
              );
              pbData[Index].Status = "Pending";
              setpbData([...pbData]);
            })
            .catch(pbErrorFunction);
        }
        setpbModalBoxVisibility(false);
        setpbPopup("DRSuccess");
        setTimeout(() => {
          setpbPopup("Close");
        }, 2000);
      })
      .catch(pbErrorFunction);
  };
  const cancelPBData = () => {
    setpbFilterOptions({ ...pbFilterKeys });
    setpbData([...pbMasterData]);
    setpbUpdate(false);
    let pbFilter = ProductionBoardFilter([...pbMasterData], pbFilterKeys);
    setpbFilterData(pbFilter);
    paginate(1, pbFilter);
    setpbAutoSave(false);
  };
  const reloadFilterOptions = (data) => {
    let tempArrReload = data;
    tempArrReload.forEach((item) => {
      if (
        pbDrpDwnOptns.BA.findIndex((BA) => {
          return BA.key == item.BA;
        }) == -1 &&
        item.BA
      ) {
        pbDrpDwnOptns.BA.push({
          key: item.BA,
          text: item.BA,
        });
      }
      if (
        pbDrpDwnOptns.Source.findIndex((Source) => {
          return Source.key == item.Source;
        }) == -1 &&
        item.Source
      ) {
        pbDrpDwnOptns.Source.push({
          key: item.Source,
          text: item.Source,
        });
      }
      if (
        pbDrpDwnOptns.Product.findIndex((Product) => {
          return Product.key == item.Product;
        }) == -1 &&
        item.Product
      ) {
        pbDrpDwnOptns.Product.push({
          key: item.Product,
          text: item.Product,
        });
      }
      if (
        pbDrpDwnOptns.Project.findIndex((Project) => {
          return Project.key == item.Project;
        }) == -1 &&
        item.Project
      ) {
        pbDrpDwnOptns.Project.push({
          key: item.Project,
          text: item.Project,
        });
      }
    });
    setpbDropDownOptions(pbDrpDwnOptns);
  };
  const dpValidationFunction = () => {
    let isError = false;

    let errorStatus = {
      Request: "",
      Requestto: "",
      Documenttype: "",
      Link: "",
    };

    if (!pbDocumentReview.Request) {
      isError = true;
      errorStatus.Request = "Please select a value for request";
    }
    if (!pbDocumentReview.Requestto) {
      isError = true;
      errorStatus.Requestto = "Please select a value for request to";
    }
    if (!pbDocumentReview.Documenttype) {
      isError = true;
      errorStatus.Documenttype = "Please select a value for document type";
    }
    if (!pbDocumentReview.Link) {
      isError = true;
      errorStatus.Link = "Please enter a value for link";
    }

    if (!isError) {
      setpbButtonLoader(true);
      savePBDRData();
    } else {
      setpbShowMessage(errorStatus);
    }
  };
  const pbErrorFunction = (error) => {
    setpbLoader(false);
    console.log(error);
    setpbPopup("Error");
    setTimeout(() => {
      setpbPopup("Close");
    }, 2000);
  };
  const SuccessPopup = () => (
    <MessageBar messageBarType={MessageBarType.success} isMultiline={false}>
      Production board has been Successfully Submitted !!!
    </MessageBar>
  );
  const DRSuccessPopup = () => (
    <MessageBar messageBarType={MessageBarType.success} isMultiline={false}>
      Document review has been Successfully Submitted !!!
    </MessageBar>
  );
  const ErrorPopup = () => (
    <MessageBar messageBarType={MessageBarType.error} isMultiline={false}>
      Something when Error, please contact system admin.
    </MessageBar>
  );

  //Onchange Function
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
  const onChangeFilter = (key, option) => {
    let tempArr = [];
    let tempDpFilterKeys = { ...pbFilterOptions };
    tempDpFilterKeys[`${key}`] = option;

    if (tempDpFilterKeys.Week == "This Week") {
      setpbWeek(Pb_WeekNumber);
      setpbYear(Pb_Year);
      tempArr = [...pbData];
    } else if (tempDpFilterKeys.Week == "Last Week") {
      cancelPBData();
      setpbWeek(Pb_LastWeekNumber);
      setpbYear(Pb_LastWeekYear);
      tempArr = [...pbLastweek];
    } else if (tempDpFilterKeys.Week == "Next Week") {
      cancelPBData();
      setpbWeek(Pb_NextWeekNumber);
      setpbYear(Pb_NextWeekYear);
      tempArr = [...pbNextweek];
    }

    if (tempDpFilterKeys.BA != "All") {
      tempArr = tempArr.filter((arr) => {
        return arr.BA == tempDpFilterKeys.BA;
      });
    }
    if (tempDpFilterKeys.Source != "All") {
      tempArr = tempArr.filter((arr) => {
        return arr.Source == tempDpFilterKeys.Source;
      });
    }
    if (tempDpFilterKeys.Product != "All") {
      tempArr = tempArr.filter((arr) => {
        return arr.Product == tempDpFilterKeys.Product;
      });
    }
    if (tempDpFilterKeys.Project != "All") {
      tempArr = tempArr.filter((arr) => {
        return arr.Project == tempDpFilterKeys.Project;
      });
    }

    setpbFilterOptions({ ...tempDpFilterKeys });
    let pbFilter = ProductionBoardFilter([...tempArr], tempDpFilterKeys);
    setpbFilterData(pbFilter);
    paginate(1, pbFilter);
  };
  const pbOnchangeItems = (RefId, key, value) => {
    let Index = pbData.findIndex((obj) => obj.RefId == RefId);
    let filIndex = pbFilterData.findIndex((obj) => obj.RefId == RefId);
    let disIndex = pbDisplayData.findIndex((obj) => obj.RefId == RefId);
    let pbBeforeData = pbData[Index];

    let pbOnchangeData = [
      {
        ID: pbBeforeData.ID,
        BA: pbBeforeData.BA,
        StartDate: pbBeforeData.StartDate,
        EndDate: pbBeforeData.EndDate,
        Source: pbBeforeData.Source,
        Project: pbBeforeData.Project,
        AnnualPlanID: pbBeforeData.AnnualPlanID,
        ProductId: pbBeforeData.ProductId,
        Product: pbBeforeData.Product,
        Title: pbBeforeData.Title,
        PlannedHours: pbBeforeData.PlannedHours,
        Monday: key == "Monday" ? value : pbBeforeData.Monday,
        Tuesday: key == "Tuesday" ? value : pbBeforeData.Tuesday,
        Wednesday: key == "Wednesday" ? value : pbBeforeData.Wednesday,
        Thursday: key == "Thursday" ? value : pbBeforeData.Thursday,
        Friday: key == "Friday" ? value : pbBeforeData.Friday,
        ActualHours: pbBeforeData.ActualHours,
        DeveloperId: pbBeforeData.DeveloperId,
        DeveloperEmail: pbBeforeData.DeveloperEmail,
        RefId: pbBeforeData.RefId,
        Week: pbBeforeData.Week,
        Year: pbBeforeData.Year,
        NotApplicable: pbBeforeData.NotApplicable,
        NotApplicableManager: pbBeforeData.NotApplicableManager,
        DeliveryPlanID: pbBeforeData.DeliveryPlanID,
        DPActualHours: pbBeforeData.DRActualHours,
        Status: pbBeforeData.Status,
      },
    ];
    pbOnchangeData[0]["ActualHours"] =
      parseInt(pbOnchangeData[0]["Monday"] ? pbOnchangeData[0]["Monday"] : 0) +
      parseInt(
        pbOnchangeData[0]["Tuesday"] ? pbOnchangeData[0]["Tuesday"] : 0
      ) +
      parseInt(
        pbOnchangeData[0]["Wednesday"] ? pbOnchangeData[0]["Wednesday"] : 0
      ) +
      parseInt(
        pbOnchangeData[0]["Thursday"] ? pbOnchangeData[0]["Thursday"] : 0
      ) +
      parseInt(pbOnchangeData[0]["Friday"] ? pbOnchangeData[0]["Friday"] : 0);

    pbData[Index] = pbOnchangeData[0];
    pbFilterData[filIndex] = pbOnchangeData[0];
    pbDisplayData[disIndex] = pbOnchangeData[0];
    setpbData([...pbData]);
    setpbFilterData([...pbFilterData]);
    setpbDisplayData([...pbDisplayData]);
  };
  const pbAddOnchange = (key, value) => {
    let tempArronchange = pbDocumentReview;
    if (key == "Request") tempArronchange.Request = value;
    else if (key == "Requestto") tempArronchange.Requestto = value;
    else if (key == "Emailcc") tempArronchange.Emailcc = value;
    //else if (key == "Project") tempArronchange.Project = value;
    else if (key == "Documenttype") tempArronchange.Documenttype = value;
    else if (key == "Link") tempArronchange.Link = value;
    else if (key == "Comments") tempArronchange.Comments = value;
    else if (key == "Confidential") tempArronchange.Confidential = value;

    console.log(tempArronchange);
    setpbDocumentReview(tempArronchange);
  };
  const paginate = (pagenumber, data) => {
    if (data.length > 0) {
      let lastIndex: number = pagenumber * totalPageItems;
      let firstIndex: number = lastIndex - totalPageItems;
      let paginatedItems = data.slice(firstIndex, lastIndex);
      currentpage = pagenumber;
      setpbDisplayData(paginatedItems);
      setpbCurrentPage(pagenumber);
    } else {
      setpbDisplayData([]);
      setpbCurrentPage(1);
    }
  };
  const ProductionBoardFilter = (data, filterValue) => {
    let tempArr = data.filter(
      (pb) => pb.NotApplicable != true && pb.NotApplicableManager != true
    );
    let tempDpFilterKeys = { ...filterValue };

    if (tempDpFilterKeys.Showonly == "Mine") {
      tempArr = tempArr.filter((arr) => {
        return arr.DeveloperEmail == loggeduseremail;
      });
    }

    if (tempDpFilterKeys.Week == "This Week") {
      tempArr = tempArr.filter((arr) => {
        let start = moment(arr.StartDate).isoWeek();
        let end = moment(arr.EndDate).isoWeek();
        let today = Pb_WeekNumber;
        return (
          today >= start && today <= end
          //(today > start && today <= end) || (today >= start && today < end)
        );
      });
    } else if (tempDpFilterKeys.Week == "Last Week") {
      tempArr = tempArr.filter((arr) => {
        let start = moment(arr.StartDate).isoWeek();
        let end = moment(arr.EndDate).isoWeek();
        let today = Pb_LastWeekNumber;
        return (
          today >= start && today <= end
          // (today > start && today <= end) || (today >= start && today < end)
        );
      });
    } else if (tempDpFilterKeys.Week == "Next Week") {
      tempArr = tempArr.filter((arr) => {
        let start = moment(arr.StartDate).isoWeek();
        let end = moment(arr.EndDate).isoWeek();
        let today = Pb_NextWeekNumber;
        return (
          today >= start && today <= end
          //(today > start && today <= end) || (today >= start && today < end)
        );
      });
    }

    return tempArr;
  };
  const disHours = () => {
    var sum: number = 0;
    let tempArr = pbFilterData.filter((arr) => {
      return arr.DeveloperEmail == loggeduseremail;
    });
    if (tempArr.length > 0) {
      tempArr.forEach((x) => {
        sum += parseInt(x.PlannedHours ? x.PlannedHours : 0);
      });
      return sum ? sum : 0;
    } else {
      return sum ? sum : 0;
    }
  };
  // Return function
  return (
    <div style={{ padding: "5px 15px" }}>
      <>
        {pbPopup == "Success"
          ? SuccessPopup()
          : pbPopup == "DRSuccess"
          ? DRSuccessPopup()
          : pbPopup == "Error"
          ? ErrorPopup()
          : ""}
      </>
      {pbLoader ? <CustomLoader /> : null}
      <div className={styles.apHeaderSection} style={{ paddingBottom: "0" }}>
        <div
          style={{
            display: "flex",
            alignItems: "center",
            justifyContent: "space-between",
            marginBottom: 20,
            color: "#2392b2",
          }}
        >
          <div className={styles.dpTitle}>
            {Ap_AnnualPlanId ? (
              <Icon
                aria-label="ChevronLeftMed"
                iconName="ChevronLeftMed"
                className={pbBigiconStyleClass.ChevronLeftMed}
                onClick={() => {
                  pbAutoSave
                    ? alertDialogforBack()
                    : navType == "AP"
                    ? props.handleclick("AnnualPlan")
                    : props.handleclick("DeliveryPlan", Ap_AnnualPlanId);
                }}
              />
            ) : null}
            <Label style={{ color: "#000", fontSize: 24, padding: 0 }}>
              Production board
            </Label>
          </div>
        </div>
        <div
          style={{
            display: "flex",
            alignItems: "center",
            justifyContent: "space-between",
            marginBottom: 20,
            color: "#2392b2",
          }}
        >
          <div
            style={{
              display: "flex",
              alignItems: "center",
              justifyContent: "center",
            }}
          >
            <div
              style={{
                display: "flex",
                alignItems: "center",
                justifyContent: "center",
              }}
              className="toggleWrapper"
            >
              {/* <Label
              className={
                !pbChecked
                  ? pblabelStyles.selectedLabel
                  : pblabelStyles.titleLabel
              }
            >
              Deliverables
            </Label>
            <Toggle
              style={{
                marginTop: 10,
                marginRight: 10,
              }}
              defaultChecked={pbChecked}
              //label={<div>Custom inline label {iconWithTooltip}</div>}
              inlineLabel
              //   onText="On"
              //   offText="Off"
              onChange={(ev) => {
                setpbChecked(!pbChecked);
              }}
            />
            <Label
              className={
                pbChecked
                  ? pblabelStyles.selectedLabel
                  : pblabelStyles.titleLabel
              }
            >
              Activity Planner
            </Label> */}
              <label
                htmlFor="toggle"
                className={styles.toggle}
                onChange={(ev) => {
                  setpbChecked(!pbChecked);
                }}
              >
                <input type="checkbox" id="toggle" />
                <span className={styles.slider}>
                  <p>Deliverables</p>
                  <p>Activity Planner</p>
                  {/* <span className={styles.sliderBar}></span> */}
                </span>
              </label>
            </div>
            <div className={pbProjectInfo}>
              <Label className={pblabelStyles.titleLabel}>Week :</Label>
              <Label
                className={pblabelStyles.labelValue}
                style={{ maxWidth: 500 }}
              >
                {pbWeek}
              </Label>
            </div>
            <div className={pbProjectInfo}>
              <Label className={pblabelStyles.titleLabel}>Year :</Label>
              <Label
                className={pblabelStyles.labelValue}
                style={{ maxWidth: 500 }}
              >
                {pbYear}
              </Label>
            </div>
            <div className={pbProjectInfo}>
              <Label className={pblabelStyles.titleLabel}>Hours :</Label>
              <Label
                className={pblabelStyles.labelValue}
                style={{ maxWidth: 500 }}
              >
                {disHours()}
              </Label>
            </div>
          </div>
          <div
            style={{
              display: "flex",
              alignItems: "center",
              justifyContent: "center",
            }}
          >
            {pbData.length > 0 && pbWeek == Pb_WeekNumber ? (
              <div>
                {pbUpdate ? (
                  <div>
                    <PrimaryButton
                      iconProps={cancelIcon}
                      text="Cancel"
                      className={pbbuttonStyleClass.buttonPrimary}
                      onClick={(_) => {
                        cancelPBData();
                      }}
                    />
                    <PrimaryButton
                      iconProps={saveIcon}
                      text="Save"
                      id="btnSave"
                      className={pbbuttonStyleClass.buttonSecondary}
                      onClick={(_) => {
                        setpbAutoSave(false);
                        savePBData();
                      }}
                    />
                  </div>
                ) : (
                  <div>
                    <PrimaryButton
                      iconProps={editIcon}
                      text="Edit"
                      className={pbbuttonStyleClass.buttonPrimary}
                      onClick={() => {
                        setpbUpdate(true);
                        setpbAutoSave(true);
                      }}
                    />
                    <PrimaryButton
                      iconProps={saveIcon}
                      text="Save"
                      disabled={true}
                      //   className={pbbuttonStyleClass.buttonSecondary}
                      onClick={(_) => {}}
                    />
                  </div>
                )}
              </div>
            ) : null}
            <Label
              onClick={() => {
                generateExcel();
              }}
              style={{
                backgroundColor: "#EBEBEB",
                padding: "7px 15px",
                cursor: "pointer",
                fontSize: "12px",
                display: "flex",
                alignItems: "center",
                justifyContent: "center",
                borderRadius: "3px",
                color: "#1D6F42",
                marginLeft: 10,
              }}
            >
              <Icon
                style={{
                  color: "#1D6F42",
                }}
                iconName="ExcelDocument"
                className={pbiconStyleClass.export}
              />
              Export as XLS
            </Label>
            {false ? (
              <Icon
                iconName="PasteAsText"
                className={pbiconStyleClass.pblink}
                onClick={() => {
                  props.handleclick("ProductionBoard", Ap_AnnualPlanId);
                }}
              />
            ) : null}
          </div>
        </div>
        <div
          style={{
            display: "flex",
            alignItems: "center",
            justifyContent: "space-between",
            // marginBottom: "15px",
            // paddingBottom: "10px",
          }}
        >
          <div className={styles.ddSection}>
            <div>
              <Label styles={pbLabelStyles}>Business area</Label>
              <Dropdown
                placeholder="Select an option"
                options={pbDropDownOptions.BA}
                selectedKey={
                  Ap_AnnualPlanId && pbFilterData.length > 0
                    ? pbFilterData[0].BA
                    : pbFilterOptions.BA
                }
                styles={
                  pbFilterOptions.BA == "All"
                    ? pbDropdownStyles
                    : pbActiveDropdownStyles
                }
                onChange={(e, option: any) => {
                  onChangeFilter("BA", option["key"]);
                }}
              />
            </div>
            <div>
              <Label styles={pbLabelStyles}>Source</Label>
              <Dropdown
                selectedKey={pbFilterOptions.Source}
                placeholder="Select an option"
                options={pbDropDownOptions.Source}
                styles={
                  pbFilterOptions.Source == "All"
                    ? pbDropdownStyles
                    : pbActiveDropdownStyles
                }
                onChange={(e, option: any) => {
                  onChangeFilter("Source", option["key"]);
                }}
              />
            </div>
            <div>
              <Label styles={pbLabelStyles}>Product</Label>
              <Dropdown
                selectedKey={
                  Ap_AnnualPlanId &&
                  pbFilterData.length > 0 &&
                  pbFilterData[0].Product
                    ? pbFilterData[0].Product
                    : pbFilterOptions.Product
                }
                placeholder="Select an option"
                options={pbDropDownOptions.Product}
                styles={
                  pbFilterOptions.Product == "All"
                    ? pbDropdownStyles
                    : pbActiveDropdownStyles
                }
                onChange={(e, option: any) => {
                  onChangeFilter("Product", option["key"]);
                }}
              />
            </div>
            <div>
              <Label styles={pbLabelStyles}>Project or task</Label>
              <Dropdown
                selectedKey={
                  Ap_AnnualPlanId && pbFilterData.length > 0
                    ? pbFilterData[0].Project
                    : pbFilterOptions.Project
                }
                placeholder="Select an option"
                options={pbDropDownOptions.Project}
                dropdownWidth={"auto"}
                styles={
                  pbFilterOptions.Project == "All"
                    ? pbDropdownStyles
                    : pbActiveDropdownStyles
                }
                onChange={(e, option: any) => {
                  onChangeFilter("Project", option["key"]);
                }}
              />
            </div>
            <div>
              <Label styles={pbLabelStyles}>Show only</Label>
              <Dropdown
                selectedKey={pbFilterOptions.Showonly}
                placeholder="Select an option"
                options={pbDropDownOptions.Showonly}
                styles={
                  pbFilterOptions.Showonly == "Mine"
                    ? pbDropdownStyles
                    : pbActiveDropdownStyles
                }
                onChange={(e, option: any) => {
                  onChangeFilter("Showonly", option["key"]);
                }}
              />
            </div>
            <div>
              <Label styles={pbLabelStyles}>Week</Label>
              <Dropdown
                selectedKey={pbFilterOptions.Week}
                placeholder="Select an option"
                options={pbDropDownOptions.Week}
                styles={
                  pbFilterOptions.Week == "This Week"
                    ? pbDropdownStyles
                    : pbActiveDropdownStyles
                }
                onChange={(e, option: any) => {
                  onChangeFilter("Week", option["key"]);
                }}
              />
            </div>

            <div>
              <div>
                <Icon
                  iconName="Refresh"
                  className={pbiconStyleClass.refresh}
                  onClick={() => {
                    if (pbAutoSave) {
                      if (
                        confirm(
                          "You have unsaved changes, are you sure you want to leave?"
                        )
                      ) {
                        setpbData([...pbMasterData]);
                        setpbFilterOptions({ ...pbFilterKeys });
                        let pbFilter = ProductionBoardFilter(
                          [...pbMasterData],
                          pbFilterKeys
                        );
                        setpbFilterData(pbFilter);
                        paginate(1, pbFilter);
                        setpbUpdate(false);
                      }
                    } else {
                      setpbData([...pbMasterData]);
                      setpbFilterOptions({ ...pbFilterKeys });
                      let pbFilter = ProductionBoardFilter(
                        [...pbMasterData],
                        pbFilterKeys
                      );
                      setpbFilterData(pbFilter);
                      paginate(1, pbFilter);
                      setpbUpdate(false);
                    }
                  }}
                />
              </div>
            </div>
          </div>
          <div
            className={pbProjectInfo}
            style={{ marginLeft: "20px", transform: "translateY(12px)" }}
          >
            <Label className={pblabelStyles.NORLabel}>
              Number of records:{" "}
              <b style={{ color: "#038387" }}>{pbFilterData.length}</b>
            </Label>
          </div>
        </div>
      </div>
      {pbChecked ? (
        <div style={{ marginTop: "10px" }}>
          <DetailsList
            items={pbDisplayData}
            columns={_dpColumns}
            setKey="set"
            layoutMode={DetailsListLayoutMode.justified}
            selectionMode={SelectionMode.none}
            styles={gridStyles}
          />
          <div
            style={{
              display: "flex",
              justifyContent: "center",
              margin: "20px 0",
            }}
          >
            {pbFilterData.length > 0 ? (
              <Pagination
                currentPage={pbcurrentPage}
                totalPages={
                  pbFilterData.length > 0
                    ? Math.ceil(pbFilterData.length / totalPageItems)
                    : 1
                }
                onChange={(page) => {
                  paginate(page, pbFilterData);
                }}
              />
            ) : (
              <div
                style={{
                  display: "flex",
                  justifyContent: "center",
                  marginTop: "15px",
                }}
              >
                <Label>No data Found !!!</Label>
              </div>
            )}
          </div>
        </div>
      ) : (
        <div
          style={{
            display: "flex",
            justifyContent: "center",
            marginTop: "15px",
          }}
        >
          <Label>No data Found !!!</Label>
        </div>
      )}

      <Modal isOpen={pbModalBoxVisibility} isBlocking={false}>
        <div style={{ padding: "30px 20px" }}>
          <div
            style={{
              fontSize: 24,
              textAlign: "center",
              color: "#2392B2",
              fontWeight: "600",
              marginBottom: "20px",
            }}
          >
            Document review
          </div>
          <div
            style={{
              display: "flex",
              alignItems: "flex-start",
              justifyContent: "flex-start",
            }}
          >
            <div>
              <Dropdown
                required={true}
                errorMessage={pbShowMessage.Request}
                label="Request"
                //selectedKey={apResponseData.product}
                placeholder="Select an option"
                options={pbModalBoxDropDownOptions.Request}
                styles={pbModalBoxDrpDwnCalloutStyles}
                onChange={(e, option: any) => {
                  pbAddOnchange("Request", option["key"]);
                }}
              />
            </div>
            <div>
              <Label
                required={true}
                style={{
                  transform: "translate(20px, 10px)",
                }}
              >
                Request to
              </Label>
              <NormalPeoplePicker
                className={pbModalBoxPP}
                // errorMessage={pbShowMessage.Requestto}
                onResolveSuggestions={GetUserDetails}
                itemLimit={1}
                onChange={(selectedUser) => {
                  selectedUser.length != 0
                    ? pbAddOnchange("Requestto", selectedUser[0]["ID"])
                    : pbAddOnchange("Requestto", "");
                }}
              />
              <Label
                style={{
                  transform: "translate(20px, 10px)",
                  color: "#a4262c",
                  fontSize: 12,
                  fontWeight: 400,
                  paddingTop: 5,
                  marginTop: -20,
                }}
              >
                {pbShowMessage.Requestto}
              </Label>
            </div>
            <div>
              <Label
                style={{
                  transform: "translate(20px, 10px)",
                }}
              >
                Email (cc)
              </Label>
              <NormalPeoplePicker
                className={pbModalBoxPP}
                onResolveSuggestions={GetUserDetails}
                itemLimit={5}
                onChange={(selectedUser) => {
                  let selectedId = selectedUser.map((su) => su["ID"]);
                  selectedUser.length != 0
                    ? pbAddOnchange("Emailcc", selectedId)
                    : pbAddOnchange("Emailcc", "");
                }}
              />
            </div>
          </div>
          <div
            style={{
              display: "flex",
              alignItems: "flex-start",
              justifyContent: "flex-start",
            }}
          >
            <div>
              <TextField
                label="Project or task"
                placeholder="Add new project"
                // errorMessage={apShowMessage.projectOrTaskError}
                defaultValue={pbDocumentReview.Project}
                disabled={true}
                styles={pbTxtBoxStyles}
                className={styles.projectField}
                onChange={(e, value: string) => {
                  //pbAddOnchange("Project", value);
                }}
              />
            </div>
            <div>
              <Dropdown
                label="Document type"
                required={true}
                errorMessage={pbShowMessage.Documenttype}
                //selectedKey={apResponseData.product}
                placeholder="Select an option"
                options={pbModalBoxDropDownOptions.Documenttype}
                styles={pbModalBoxDrpDwnCalloutStyles}
                onChange={(e, option: any) => {
                  pbAddOnchange("Documenttype", option["key"]);
                }}
              />
            </div>
            <div>
              <TextField
                label="Link"
                placeholder="Add link"
                errorMessage={pbShowMessage.Link}
                // defaultValue={apResponseData.projectOrTask}
                required={true}
                styles={pbTxtBoxStyles}
                onChange={(e, value: string) => {
                  pbAddOnchange("Link", value);
                }}
              />
            </div>
          </div>
          <div
            style={{
              display: "flex",
              alignItems: "flex-start",
              justifyContent: "flex-start",
            }}
          >
            <div>
              <TextField
                label="Comments"
                placeholder="Add Comments"
                // errorMessage={apShowMessage.projectOrTaskError}
                // defaultValue={apResponseData.projectOrTask}
                // required={true}
                multiline
                rows={5}
                resizable={false}
                styles={pbMultiTxtBoxStyles}
                onChange={(e, value: string) => {
                  pbAddOnchange("Comments", value);
                }}
              />
            </div>
            <div
              style={{
                marginTop: 30,
                marginLeft: 20,
                position: "relative",
              }}
            >
              <Toggle
                //defaultChecked={pbChecked}
                label={
                  <div
                    style={{
                      position: "absolute",
                      left: "0",
                      top: "0",
                      width: "200px",
                    }}
                  >
                    Confidential
                  </div>
                }
                inlineLabel
                style={{ transform: "translateX(100px)" }}
                onChange={(ev) => {
                  pbAddOnchange("Confidential", !pbDocumentReview.Confidential);
                }}
              />
            </div>
          </div>
          <div className={styles.apModalBoxButtonSection}>
            <button
              className={styles.apModalBoxSubmitBtn}
              onClick={(_) => {
                dpValidationFunction();
              }}
              style={{ display: "flex" }}
            >
              {pbButtonLoader ? (
                <Spinner />
              ) : (
                <span>
                  <Icon
                    iconName="Save"
                    style={{ position: "relative", top: 3, left: -8 }}
                  />
                  {"Submit"}
                </span>
              )}
            </button>
            <button
              className={styles.apModalBoxBackBtn}
              onClick={(_) => {
                setpbModalBoxVisibility(false);
              }}
            >
              <span>
                {" "}
                <Icon
                  iconName="Cancel"
                  style={{ position: "relative", top: 3, left: -8 }}
                />
                Close
              </span>
            </button>
          </div>
        </div>
      </Modal>
    </div>
  );
}

export default ProductionBoard;
