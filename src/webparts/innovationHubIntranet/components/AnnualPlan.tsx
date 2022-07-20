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
  SearchBox,
  ISearchBoxStyles,
  Dropdown,
  IDropdownStyles,
  NormalPeoplePicker,
  Persona,
  PersonaPresence,
  PersonaSize,
  Modal,
  DatePicker,
  IDatePickerStyles,
  Checkbox,
  ICheckboxStyles,
  TextField,
  ITextFieldStyles,
  MessageBar,
  MessageBarType,
  Spinner,
  SpinnerSize,
  TooltipHost,
  TooltipOverflowMode,
} from "@fluentui/react";
import Pagination from "office-ui-fabric-react-pagination";
import * as Excel from "exceljs/dist/exceljs.min.js";
import * as FileSaver from "file-saver";
import "../ExternalRef/styleSheets/ProdStyles.css";
import { mergeStyles, mergeStyleSets } from "@fluentui/react/lib/Styling";
import styles from "./InnovationHubIntranet.module.scss";
import CustomLoader from "./CustomLoader";

const AnnualPlan = (props: any) => {
  const sharepointWeb = Web(props.URL);
  const ListNameURL = props.WeblistURL;
  let currentpage = 1;
  let totalPageItems = 10;
  const apAllitems = [];
  const apMasterProductCollection = [];
  const allPeoples = [];
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
  const apColumns = props.isAdmin
    ? [
        {
          key: "BAacronyms",
          name: "BA",
          fieldName: "BAacronyms",
          minWidth: 30,
          maxWidth: 30,
        },
        {
          key: "Term",
          name: "Term",
          fieldName: "Term",
          minWidth: 40,
          maxWidth: 40,
        },
        {
          key: "Hours",
          name: "Hours",
          fieldName: "Hours",
          minWidth: 40,
          maxWidth: 40,
        },
        {
          key: "StartDate",
          name: "Start date",
          fieldName: "StartDate",
          minWidth: 65,
          maxWidth: 65,
        },
        {
          key: "EndDate",
          name: "End date",
          fieldName: "EndDate",
          minWidth: 70,
          maxWidth: 70,
        },
        {
          key: "Product",
          name: "Product or solution",
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
          key: "TypeOfProject",
          name: "TOD",
          fieldName: "TypeOfProject",
          minWidth: 30,
          maxWidth: 30,
        },
        {
          key: "ProjectOrTask",
          name: "Project or task",
          fieldName: "ProjectOrTask",
          minWidth: 300,
          maxWidth: 300,

          onRender: (item) => (
            <>
              <TooltipHost
                id={item.ID}
                content={item.ProjectOrTask}
                overflowMode={TooltipOverflowMode.Parent}
              >
                <span aria-describedby={item.ID}>{item.ProjectOrTask}</span>
              </TooltipHost>
            </>
          ),
        },
        {
          key: "DP/AP",
          name: "DP",
          fieldName: "DPAP",
          minWidth: 30,
          maxWidth: 30,

          onRender: (item) => (
            <>
              <Icon
                style={{
                  // fontSize: 24,
                  // marginTop: -6,
                  marginLeft: 0,
                }}
                iconName="Link12"
                className={apIconStyleClass.link}
                onClick={() => {
                  props.handleclick("DeliveryPlan", item.ID, "AP");
                }}
              />
            </>
          ),
        },
        {
          key: "PB",
          name: "PB",
          fieldName: "PB",
          minWidth: 30,
          maxWidth: 30,

          onRender: (item) => (
            <>
              <Icon
                iconName="Link12"
                className={apIconStyleClass.link}
                onClick={() => {
                  props.handleclick("ProductionBoard", item.ID, "AP");
                }}
              />
            </>
          ),
        },
        {
          key: "Status",
          name: "Status",
          fieldName: "Status",
          minWidth: 120,
          maxWidth: 120,
        },
        {
          key: "PM",
          name: "PM",
          fieldName: "PM",
          minWidth: 50,
          maxWidth: 50,
        },
        {
          key: "D",
          name: "D",
          fieldName: "D",
          minWidth: 50,
          maxWidth: 50,
        },
        {
          key: "Actions",
          name: "Actions",
          fieldName: "Actions",
          minWidth: 90,
          maxWidth: 90,

          onRender: (item) => (
            <>
              <Icon
                iconName="Edit"
                className={apIconStyleClass.edit}
                onClick={() => {
                  let filteredArr = apMasterData.filter((data) => {
                    return data.ID == item.ID;
                  });
                  setApResponseData({
                    ID: filteredArr[0].ID,
                    businessArea: filteredArr[0].BusinessArea,
                    typeOfProject: filteredArr[0].TypeOfProject,
                    term: filteredArr[0].Term,
                    product: filteredArr[0].Product,
                    startDate: filteredArr[0].StartDate
                      ? new Date(filteredArr[0].DefaultStartDate)
                      : null,
                    endDate: filteredArr[0].EndDate
                      ? new Date(filteredArr[0].DefaultEndDate)
                      : null,
                    projectOrTask: filteredArr[0].ProjectOrTask,
                    year: filteredArr[0].Year,
                    manager: filteredArr[0].PMName.id,
                    developer:
                      filteredArr[0].DNames.length > 0
                        ? filteredArr[0].DNames[0].id
                        : "",
                  });
                  setApModalBoxVisibility({
                    condition: true,
                    action: "Update",
                    selectedItem: filteredArr,
                  });
                }}
              />
              <Icon
                iconName="Delete"
                className={apIconStyleClass.delete}
                onClick={() => {
                  setApDeletePopup({ condition: true, targetId: item.ID });
                }}
              />
            </>
          ),
        },
      ]
    : [
        {
          key: "BAacronyms",
          name: "BA",
          fieldName: "BAacronyms",
          minWidth: 40,
          maxWidth: 40,
        },
        {
          key: "Term",
          name: "T",
          fieldName: "Term",
          minWidth: 20,
          maxWidth: 20,
        },
        {
          key: "Hours",
          name: "Hours",
          fieldName: "Hours",
          minWidth: 50,
          maxWidth: 50,
        },
        {
          key: "StartDate",
          name: "Start date",
          fieldName: "StartDate",
          minWidth: 100,
          maxWidth: 100,
        },
        {
          key: "EndDate",
          name: "End date",
          fieldName: "EndDate",
          minWidth: 120,
          maxWidth: 120,
        },
        {
          key: "Product",
          name: "Product or solution",
          fieldName: "Product",
          minWidth: 150,
          maxWidth: 150,
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
          key: "TypeOfProject",
          name: "TOD",
          fieldName: "TypeOfProject",
          minWidth: 40,
          maxWidth: 40,
        },
        {
          key: "ProjectOrTask",
          name: "Project or task",
          fieldName: "ProjectOrTask",
          minWidth: 245,
          maxWidth: 245,

          onRender: (item) => (
            <>
              <TooltipHost
                id={item.ID}
                content={item.ProjectOrTask}
                overflowMode={TooltipOverflowMode.Parent}
              >
                <span aria-describedby={item.ID}>{item.ProjectOrTask}</span>
              </TooltipHost>
            </>
          ),
        },
        {
          key: "DP/AP",
          name: "DP/AP",
          fieldName: "DPAP",
          minWidth: 20,
          maxWidth: 20,

          onRender: (item) => (
            <>
              <Icon
                iconName="Link12"
                className={apIconStyleClass.link}
                onClick={() => {
                  props.handleclick("DeliveryPlan", item.ID, "AP");
                }}
              />
            </>
          ),
        },
        {
          key: "PB",
          name: "PB",
          fieldName: "PB",
          minWidth: 20,
          maxWidth: 20,

          onRender: (item) => (
            <>
              <Icon
                iconName="Link12"
                className={apIconStyleClass.link}
                onClick={() => {
                  props.handleclick("ProductionBoard", item.ID, "AP");
                }}
              />
            </>
          ),
        },
        {
          key: "Status",
          name: "Status",
          fieldName: "Status",
          minWidth: 120,
          maxWidth: 120,
        },
        {
          key: "PM",
          name: "PM",
          fieldName: "PM",
          minWidth: 30,
          maxWidth: 30,
        },
        {
          key: "D",
          name: "D",
          fieldName: "D",
          minWidth: 30,
          maxWidth: 30,
        },
      ];
  const apDrpDwnOptns = {
    baOptns: [{ key: "All", text: "All" }],
    todOptns: [{ key: "All", text: "All" }],
    potOptns: [{ key: "All", text: "All" }],
    managerOptns: [{ key: "All", text: "All" }],
    developerOptns: [{ key: "All", text: "All" }],
    termOptns: [{ key: "All", text: "All" }],
  };
  const apModalBoxDrpDwnOptns = {
    baOptns: [],
    todOptns: [],
    potOptns: [],
    managerOptns: [],
    developerOptns: [],
    termOptns: [],
    productOptns: [],
    yearOptns: [],
  };
  const apFilterKeys = {
    ProjectOrTaskSearch: "",
    BusinessArea: "All",
    TypeOfProject: "All",
    ProjectOrTask: "All",
    PM: "All",
    D: "All",
    Term: "All",
  };
  const responseData = {
    ID: 0,
    businessArea: "",
    typeOfProject: "",
    term: "",
    product: "",
    startDate: new Date(),
    endDate: new Date(),
    projectOrTask: "",
    year: "",
    manager: "",
    developer: "",
  };
  const apErrorStatus = {
    businessAreaError: "",
    projectOrTaskError: "",
  };

  //StylesStart

  const gridStyles: Partial<IDetailsListStyles> = {
    root: {
      // overflowX: "scroll",
      selectors: {
        "& [role=grid]": {
          display: "flex",
          flexDirection: "column",
          alignItems: "start",
          // height: "60vh",
          ".ms-DetailsRow-cell": {
            height: 42,
            // padding: "6px 12px",
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
  const apLabelStyles: Partial<ILabelStyles> = {
    root: {
      width: 150,
      marginRight: 10,
      fontSize: "13px",
      color: "#323130",
    },
  };
  const apSearchBoxStyles: Partial<ISearchBoxStyles> = {
    root: {
      width: 186,
      marginRight: "15px",
      backgroundColor: "#F5F5F7",
      outline: "none",
      border: "1px solid #E8E8EA",
      borderRadius: "4px",
    },
    icon: { fontSize: 14, color: "#000" },
  };
  const apActiveSearchBoxStyles: Partial<ISearchBoxStyles> = {
    root: {
      width: 186,
      marginRight: "15px",
      backgroundColor: "#F5F5F7",
      outline: "none",
      color: "#038387",
      border: "2px solid #038387",
      borderRadius: "4px",
    },
    icon: { fontSize: 14, color: "#038387" },
  };
  const apDropdownStyles: Partial<IDropdownStyles> = {
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
  const apActiveDropdownStyles: Partial<IDropdownStyles> = {
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
  const apModalBoxDropdownStyles: Partial<IDropdownStyles> = {
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
  const apModalBoxDrpDwnCalloutStyles: Partial<IDropdownStyles> = {
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
  const apTxtBoxStyles: Partial<ITextFieldStyles> = {
    root: { width: "850px", margin: "10px 20px", borderRadius: "4px" },
    field: { fontSize: 12, color: "#000" },
  };
  const apModalBoxDatePickerStyles: Partial<IDatePickerStyles> = {
    root: {
      width: "300px",
      margin: "10px 20px",
      borderRadius: "4px",
    },
    icon: {
      fontSize: "17px",
      color: "#000",
      fontWeight: "bold",
    },
  };
  const apModalBoxCheckBoxStyles: Partial<ICheckboxStyles> = {
    root: { marginTop: "46px", transform: "translateX(-26px)" },
    label: { fontWeight: "600" },
  };
  const apModalBoxPP = mergeStyles({
    width: "300px",
    margin: "10px 20px",
  });
  const apIconStyle = mergeStyles({
    fontSize: 17,
    height: 14,
    width: 17,
    cursor: "pointer",
  });
  const apIconStyleClass = mergeStyleSets({
    link: [{ color: "#2392B2", margin: "0" }, apIconStyle],
    delete: [{ color: "#CB1E06", margin: "0 7px " }, apIconStyle],
    edit: [{ color: "#2392B2", margin: "0 7px 0 0" }, apIconStyle],
    refresh: [
      {
        color: "white",
        fontSize: "18px",
        height: 20,
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
  const apStatusStyle = mergeStyles({
    textAlign: "center",
    borderRadius: "25px",
  });
  const apStatusStyleClass = mergeStyleSets({
    completed: [
      {
        fontWeight: "600",
        padding: "3px",
        color: "#438700",
        backgroundColor: "#D9FFB3",
      },
      apStatusStyle,
    ],
    scheduled: [
      {
        fontWeight: "600",
        padding: "3px",
        color: "#9f6700",
        backgroundColor: "#FFDB99",
      },
      apStatusStyle,
    ],
    onSchedule: [
      {
        fontWeight: "600",
        padding: "3px",
        color: "#B3B300 ",
        backgroundColor: "#FFFFB3",
      },
      apStatusStyle,
    ],
    behindScheduled: [
      {
        fontWeight: "600",
        padding: "3px",
        color: "#FF0000",
        backgroundColor: "#FFB3B3",
      },
      apStatusStyle,
    ],
  });

  //stylesEnd

  const [apReRender, setapReRender] = useState(true);
  const [apMasterData, setApMasterData] = useState(apAllitems);
  const [apData, setApData] = useState(apAllitems);
  const [displayData, setdisplayData] = useState(apAllitems);
  const [apResponseData, setApResponseData] = useState(responseData);
  const [apMasterProducts, setApMasterProducts] = useState(
    apMasterProductCollection
  );
  const [peopleList, setPeopleList] = useState(allPeoples);
  const [apDropDownOptions, setApDropDownOptions] = useState(apDrpDwnOptns);
  const [apModalBoxDropDownOptions, setApModalBoxDropDownOptions] = useState(
    apModalBoxDrpDwnOptns
  );
  const [apFilterOptions, setApFilterOptions] = useState(apFilterKeys);
  const [apModalBoxVisibility, setApModalBoxVisibility] = useState({
    condition: false,
    action: "",
    selectedItem: [],
  });
  const [apDeletePopup, setApDeletePopup] = useState({
    condition: false,
    targetId: 0,
  });
  const [apModelBoxDrpDwnToTxtBox, setApModelBoxDrpDwnToTxtBox] =
    useState(false);
  const [apcurrentPage, setApCurrentPage] = useState(currentpage);
  const [apShowMessage, setApShowMessage] = useState(apErrorStatus);
  const [apPopup, setApPopup] = useState("");
  const [apStartUpLoader, setApStartUpLoader] = useState(true);
  const [apOnSubmitLoader, setApOnSubmitLoader] = useState(false);
  const [apOnDeleteLoader, setApOnDeleteLoader] = useState(false);

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
  const getApData = () => {
    sharepointWeb.lists
      .getByTitle(ListNameURL)
      .items.select(
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
      .top(5000)
      .orderBy("Modified", false)
      .get()
      .then((items) => {
        items.forEach((item: any, index: number) => {
          allitemsArrayFormatter(item, apAllitems);
        });
        filterKeys(items);
        setApMasterData(apAllitems);
        // setApData(apAllitems);
        setApStartUpLoader(false);
        paginate(1);
      })
      .catch(apErrorFunction);
  };
  const allitemsArrayFormatter = (item, allItems) => {
    let apDevelopersNames = [];
    item.ProjectLeadId != null
      ? item.ProjectLead.forEach((dev) => {
          apDevelopersNames.push({
            name: dev.Title,
            id: dev.Id,
            email: dev.EMail,
            userDetails: peopleList.filter((people) => {
              return people.ID == dev.Id;
            })[0],
          });
        })
      : (apDevelopersNames = null);

    let tempManagerInitial =
      item.ProjectOwnerId != null ? initial(item.ProjectOwner.Title) : "";
    let tempDeveloperInitial =
      item.ProjectLeadId != null ? initial(item.ProjectLead[0].Title) : "";

    allItems.push({
      ID: item.Id ? item.Id : "",
      Hours: item.AllocatedHours ? item.AllocatedHours : "",
      DefaultStartDate: item.StartDate ? item.StartDate : "",
      StartDate: item.StartDate
        ? moment(item.StartDate).format("DD/MM/YYYY")
        : "",
      DefaultEndDate: item.PlannedEndDate ? item.PlannedEndDate : "",
      EndDate: item.PlannedEndDate
        ? moment(item.PlannedEndDate).format("DD/MM/YYYY")
        : "",
      Product:
        item.Master_x0020_ProjectId != null
          ? item.Master_x0020_Project.Title
          : "",
      TypeOfProject: item.ProjectType ? item.ProjectType : "",
      Year: item.Year ? item.Year : "",
      Term: item.Term ? item.Term : "",
      BusinessArea: item.BusinessArea ? item.BusinessArea : "",
      BAacronyms: item.BA_x0020_acronyms ? item.BA_x0020_acronyms : "",
      ProjectOrTask: item.Title ? item.Title : "",
      Status: item.Status ? (
        <>
          {item.Status == "Completed" ? (
            <div className={apStatusStyleClass.completed}>{item.Status}</div>
          ) : item.Status == "Scheduled" ? (
            <div className={apStatusStyleClass.scheduled}>{item.Status}</div>
          ) : item.Status == "On schedule" ? (
            <div className={apStatusStyleClass.onSchedule}>{item.Status}</div>
          ) : item.Status == "Behind schedule" ? (
            <div className={apStatusStyleClass.behindScheduled}>
              {item.Status}
            </div>
          ) : (
            ""
          )}
        </>
      ) : (
        ""
      ),
      StatusStage: item.Status,
      DPAP: "",
      PMName:
        item.ProjectOwnerId != null
          ? {
              name: item.ProjectOwner.Title,
              id: item.ProjectOwner.Id,
              email: item.ProjectOwner.EMail,
              userDetails: peopleList.filter((people) => {
                return people.ID == item.ProjectOwner.Id;
              })[0],
            }
          : "",
      PM:
        item.ProjectOwnerId != null ? (
          <>
            {
              <div
                title={item.ProjectOwner.Title}
                style={{
                  display: "flex",
                  alignItems: "center",
                  justifyContent: "flex-start",
                  cursor: "pointer",
                  // marginLeft: "14px",
                }}
              >
                <Persona
                  size={PersonaSize.size24}
                  presence={PersonaPresence.none}
                  imageUrl={
                    "/_layouts/15/userphoto.aspx?size=S&username=" +
                    `${item.ProjectOwner.EMail}`
                  }
                  imageAlt={tempManagerInitial}
                  imageInitials={tempManagerInitial}
                />
              </div>
            }
          </>
        ) : (
          ""
        ),
      D:
        item.ProjectLeadId != null ? (
          <>
            {
              <div
                title={item.ProjectLead[0].Title}
                style={{
                  display: "flex",
                  alignItems: "center",
                  justifyContent: "flex-start",
                  cursor: "pointer",
                  // marginLeft: "14px",
                }}
              >
                <Persona
                  showOverflowTooltip
                  size={PersonaSize.size24}
                  presence={PersonaPresence.none}
                  showInitialsUntilImageLoads={true}
                  imageUrl={
                    "/_layouts/15/userphoto.aspx?size=S&username=" +
                    `${item.ProjectLead[0].EMail}`
                  }
                  imageAlt={tempDeveloperInitial}
                  imageInitials={tempDeveloperInitial}
                />
              </div>
            }
          </>
        ) : (
          ""
        ),
      DNames: item.ProjectLeadId != null ? [...apDevelopersNames] : [],
      Actions: "",
    });

    return allItems;
  };
  const getAllOptions = () => {
    const _sortFilterKeys = (a, b) => {
      if (a.text.toLowerCase() < b.text.toLowerCase()) {
        return -1;
      }
      if (a.text.toLowerCase() > b.text.toLowerCase()) {
        return 1;
      }
      return 0;
    };

    //Product Choices
    sharepointWeb.lists
      .getByTitle("Master Product List")
      .items()
      .then((allProducts) => {
        allProducts.forEach((product) => {
          if (product != null) {
            if (
              apModalBoxDrpDwnOptns.productOptns.findIndex((productOptn) => {
                return productOptn.key == product.Title;
              }) == -1
            ) {
              apModalBoxDrpDwnOptns.productOptns.push({
                key: product.Title,
                text: product.Title,
              });
              apMasterProductCollection.push({
                productName: product.Title,
                ProductId: product.Id,
              });
            }
          }
        });
      })
      .then(() => {
        apModalBoxDrpDwnOptns.productOptns.sort(_sortFilterKeys);
      })
      .catch(apErrorFunction);

    //Business Area Choices
    sharepointWeb.lists
      .getByTitle(ListNameURL)
      .fields.getByInternalNameOrTitle("BusinessArea")()
      .then((response) => {
        response["Choices"].forEach((choice) => {
          if (choice != null) {
            if (
              apModalBoxDrpDwnOptns.baOptns.findIndex((baOptn) => {
                return baOptn.key == choice;
              }) == -1
            ) {
              apModalBoxDrpDwnOptns.baOptns.push({
                key: choice,
                text: choice,
              });
            }
          }
        });
      })
      .then(() => {
        apModalBoxDrpDwnOptns.baOptns.sort(_sortFilterKeys);
      })
      .catch(apErrorFunction);
    //Type of Deliverable Choices
    sharepointWeb.lists
      .getByTitle(ListNameURL)
      .fields.getByInternalNameOrTitle("ProjectType")()
      .then((response) => {
        response["Choices"].forEach((choice) => {
          if (choice != null) {
            if (
              apModalBoxDrpDwnOptns.todOptns.findIndex((todOptn) => {
                return todOptn.key == choice;
              }) == -1
            ) {
              apModalBoxDrpDwnOptns.todOptns.push({
                key: choice,
                text: choice,
              });
            }
          }
        });
      })
      .then(() => {
        apModalBoxDrpDwnOptns.todOptns.sort(_sortFilterKeys);
      })
      .catch(apErrorFunction);
    //Year Choices
    for (let year = 2010; year <= 2030; year++) {
      apModalBoxDrpDwnOptns.yearOptns.push({
        key: year,
        text: year,
      });
    }
    //Term Choices
    sharepointWeb.lists
      .getByTitle(ListNameURL)
      .fields.getByInternalNameOrTitle("Term")()
      .then((response) => {
        apModalBoxDrpDwnOptns.termOptns = [];
        //response["Choices"].forEach((choice) => {
        ["1", "2", "3", "4"].forEach((choice) => {
          if (choice != null) {
            if (
              apModalBoxDrpDwnOptns.termOptns.findIndex((termOptn) => {
                return termOptn.key == choice;
              }) == -1
            ) {
              apModalBoxDrpDwnOptns.termOptns.push({
                key: choice,
                text: choice,
              });
            }
          }
        });
      })
      .then(() => {
        setApMasterProducts(apMasterProductCollection);
        setApModalBoxDropDownOptions(apModalBoxDrpDwnOptns);
      })
      .catch(apErrorFunction);
  };
  const apErrorFunction = (error) => {
    setApStartUpLoader(false);
    console.log(error);
    setApPopup("Error");
    setTimeout(() => {
      setApPopup("Close");
    }, 2000);
  };
  const initial = (userName) => {
    let name = userName;
    let nameArr = name.split(" ");
    let initialStr = nameArr[0][0] + nameArr[nameArr.length - 1][0];
    return initialStr;
  };
  const filterKeys = (items) => {
    items.forEach((item) => {
      if (
        apDrpDwnOptns.baOptns.findIndex((baOptn) => {
          return baOptn.key == item.BusinessArea;
        }) == -1 &&
        item.BusinessArea
      ) {
        apDrpDwnOptns.baOptns.push({
          key: item.BusinessArea,
          text: item.BusinessArea,
        });
      }

      if (
        apDrpDwnOptns.todOptns.findIndex((todOptn) => {
          return todOptn.key == item.ProjectType;
        }) == -1 &&
        item.ProjectType
      ) {
        apDrpDwnOptns.todOptns.push({
          key: item.ProjectType,
          text: item.ProjectType,
        });
      }

      if (
        apDrpDwnOptns.potOptns.findIndex((potOptn) => {
          return potOptn.key == item.Title;
        }) == -1 &&
        item.Title
      ) {
        apDrpDwnOptns.potOptns.push({
          key: item.Title,
          text: item.Title,
        });
        apModalBoxDrpDwnOptns.potOptns.push({
          key: item.Title,
          text: item.Title,
        });
      }

      let tempmanager =
        item.ProjectOwnerId != null ? item.ProjectOwner.Title : null;
      if (
        apDrpDwnOptns.managerOptns.findIndex((managerOptn) => {
          return managerOptn.key == tempmanager;
        }) == -1 &&
        tempmanager
      ) {
        apDrpDwnOptns.managerOptns.push({
          key: tempmanager,
          text: tempmanager,
        });
        // apModalBoxDrpDwnOptns.managerOptns.push({
        //   key: tempmanager,
        //   text: tempmanager,
        // });
      }

      let tempdevelopers = [];
      if (item.ProjectLeadId != null) {
        item.ProjectLead.forEach((dev) => {
          tempdevelopers.push(dev.Title);
        });

        tempdevelopers.forEach((tempdev) => {
          if (
            apDrpDwnOptns.developerOptns.findIndex((developerOptn) => {
              return developerOptn.key == tempdev;
            }) == -1 &&
            tempdev
          ) {
            apDrpDwnOptns.developerOptns.push({
              key: tempdev,
              text: tempdev,
            });
            // apModalBoxDrpDwnOptns.developerOptns.push({
            //   key: tempdev,
            //   text: tempdev,
            // });
          }
        });
      }

      if (
        apDrpDwnOptns.termOptns.findIndex((termOptn) => {
          return termOptn.key == item.Term;
        }) == -1 &&
        item.Term
      ) {
        apDrpDwnOptns.termOptns.push({
          key: item.Term,
          text: item.Term,
        });
      }
    });

    sortingFilterKeys(apDrpDwnOptns, apModalBoxDrpDwnOptns);

    setApDropDownOptions(apDrpDwnOptns);
    setApModalBoxDropDownOptions(apModalBoxDrpDwnOptns);
  };
  const filterKeysAfterModified = (items) => {
    items.forEach((item) => {
      if (
        apDrpDwnOptns.baOptns.findIndex((baOptn) => {
          return baOptn.key == item.BusinessArea;
        }) == -1 &&
        item.BusinessArea
      ) {
        apDrpDwnOptns.baOptns.push({
          key: item.BusinessArea,
          text: item.BusinessArea,
        });
      }

      if (
        apDrpDwnOptns.todOptns.findIndex((todOptn) => {
          return todOptn.key == item.TypeOfProject;
        }) == -1 &&
        item.TypeOfProject
      ) {
        apDrpDwnOptns.todOptns.push({
          key: item.TypeOfProject,
          text: item.TypeOfProject,
        });
      }

      if (
        apDrpDwnOptns.potOptns.findIndex((potOptn) => {
          return potOptn.key == item.ProjectOrTask;
        }) == -1 &&
        item.ProjectOrTask
      ) {
        apDrpDwnOptns.potOptns.push({
          key: item.ProjectOrTask,
          text: item.ProjectOrTask,
        });
        apModalBoxDrpDwnOptns.potOptns.push({
          key: item.ProjectOrTask,
          text: item.ProjectOrTask,
        });
      }

      let tempmanager = item.PMName != null ? item.PMName.name : null;
      if (
        apDrpDwnOptns.managerOptns.findIndex((managerOptn) => {
          return managerOptn.key == tempmanager;
        }) == -1 &&
        tempmanager
      ) {
        apDrpDwnOptns.managerOptns.push({
          key: tempmanager,
          text: tempmanager,
        });
        // apModalBoxDrpDwnOptns.managerOptns.push({
        //   key: tempmanager,
        //   text: tempmanager,
        // });
      }

      let tempdevelopers = [];
      if (item.DNames.length > 0) {
        item.DNames.forEach((dev) => {
          tempdevelopers.push(dev.name);
        });

        tempdevelopers.forEach((tempdev) => {
          if (
            apDrpDwnOptns.developerOptns.findIndex((developerOptn) => {
              return developerOptn.key == tempdev;
            }) == -1 &&
            tempdev != null
          ) {
            apDrpDwnOptns.developerOptns.push({
              key: tempdev,
              text: tempdev,
            });
            // apModalBoxDrpDwnOptns.developerOptns.push({
            //   key: tempdev,
            //   text: tempdev,
            // });
          }
        });
      }

      if (
        apDrpDwnOptns.termOptns.findIndex((termOptn) => {
          return termOptn.key == item.Term;
        }) == -1 &&
        item.Term
      ) {
        apDrpDwnOptns.termOptns.push({
          key: item.Term,
          text: item.Term,
        });
      }
    });

    sortingFilterKeys(apDrpDwnOptns, apModalBoxDrpDwnOptns);

    setApDropDownOptions(apDrpDwnOptns);
    let tempArr = apModalBoxDropDownOptions;
    tempArr.potOptns = apModalBoxDrpDwnOptns.potOptns;
    setApModalBoxDropDownOptions(tempArr);
  };
  const sortingFilterKeys = (apDrpDwnOptns, apModalBoxDrpDwnOptns) => {
    const sortFilterKeys = (a, b) => {
      // if (a.text.toLowerCase() < b.text.toLowerCase()) {
      //   return -1;
      // }
      // if (a.text.toLowerCase() > b.text.toLowerCase()) {
      //   return 1;
      // }
      if (a.text < b.text) {
        return -1;
      }
      if (a.text > b.text) {
        return 1;
      }
      return 0;
    };

    apDrpDwnOptns.baOptns.shift();
    apDrpDwnOptns.baOptns.sort(sortFilterKeys);
    apDrpDwnOptns.baOptns.unshift({ key: "All", text: "All" });

    apDrpDwnOptns.todOptns.shift();
    apDrpDwnOptns.todOptns.sort(sortFilterKeys);
    apDrpDwnOptns.todOptns.unshift({ key: "All", text: "All" });

    apDrpDwnOptns.potOptns.shift();
    apDrpDwnOptns.potOptns.sort(sortFilterKeys);
    apDrpDwnOptns.potOptns.unshift({ key: "All", text: "All" });

    apModalBoxDrpDwnOptns.potOptns.sort(sortFilterKeys);

    if (
      apDrpDwnOptns.managerOptns.some((managerOptn) => {
        return (
          managerOptn.text.toLowerCase() ==
          props.spcontext.pageContext.user.displayName.toLowerCase()
        );
      })
    ) {
      apDrpDwnOptns.managerOptns.shift();
      let loginUserIndex = apDrpDwnOptns.managerOptns.findIndex((user) => {
        return (
          user.text.toLowerCase() ==
          props.spcontext.pageContext.user.displayName.toLowerCase()
        );
      });
      let loginUserData = apDrpDwnOptns.managerOptns.splice(loginUserIndex, 1);

      apDrpDwnOptns.managerOptns.sort(sortFilterKeys);
      apDrpDwnOptns.managerOptns.unshift(loginUserData[0]);
      apDrpDwnOptns.managerOptns.unshift({ key: "All", text: "All" });
    } else {
      apDrpDwnOptns.managerOptns.shift();
      apDrpDwnOptns.managerOptns.sort(sortFilterKeys);
      apDrpDwnOptns.managerOptns.unshift({ key: "All", text: "All" });
    }

    if (
      apDrpDwnOptns.developerOptns.some((developerOptn) => {
        return (
          developerOptn.text.toLowerCase() ==
          props.spcontext.pageContext.user.displayName.toLowerCase()
        );
      })
    ) {
      apDrpDwnOptns.developerOptns.shift();
      let loginUserIndex = apDrpDwnOptns.developerOptns.findIndex((user) => {
        return (
          user.text.toLowerCase() ==
          props.spcontext.pageContext.user.displayName.toLowerCase()
        );
      });
      let loginUserData = apDrpDwnOptns.developerOptns.splice(
        loginUserIndex,
        1
      );
      apDrpDwnOptns.developerOptns.sort(sortFilterKeys);
      apDrpDwnOptns.developerOptns.unshift(loginUserData[0]);
      apDrpDwnOptns.developerOptns.unshift({ key: "All", text: "All" });
    } else {
      apDrpDwnOptns.developerOptns.shift();
      apDrpDwnOptns.developerOptns.sort(sortFilterKeys);
      apDrpDwnOptns.developerOptns.unshift({ key: "All", text: "All" });
    }

    apDrpDwnOptns.termOptns.shift();
    apDrpDwnOptns.termOptns.sort(sortFilterKeys);
    apDrpDwnOptns.termOptns.unshift({ key: "All", text: "All" });
  };
  const listFilter = (key, option) => {
    let tempArr = [...apMasterData];
    let tempApFilterKeys = { ...apFilterOptions };
    tempApFilterKeys[`${key}`] = option;

    if (tempApFilterKeys.ProjectOrTaskSearch) {
      tempArr = tempArr.filter((arr) => {
        return arr.ProjectOrTask.toLowerCase().includes(
          tempApFilterKeys.ProjectOrTaskSearch.toLowerCase()
        );
      });
    }

    if (tempApFilterKeys.BusinessArea != "All") {
      tempArr = tempArr.filter((arr) => {
        return arr.BusinessArea == tempApFilterKeys.BusinessArea;
      });
    }
    if (tempApFilterKeys.TypeOfProject != "All") {
      tempArr = tempArr.filter((arr) => {
        return arr.TypeOfProject == tempApFilterKeys.TypeOfProject;
      });
    }
    if (tempApFilterKeys.ProjectOrTask != "All") {
      tempArr = tempArr.filter((arr) => {
        return arr.ProjectOrTask == tempApFilterKeys.ProjectOrTask;
      });
    }
    if (tempApFilterKeys.PM != "All") {
      tempArr = tempArr.filter((arr) => {
        return arr.PMName.name == tempApFilterKeys.PM;
      });
    }
    if (tempApFilterKeys.D != "All") {
      let devArr = [];
      tempArr.forEach((arr) => {
        if (arr.DNames.length != 0) {
          if (arr.DNames.some((DName) => DName.name == tempApFilterKeys.D)) {
            devArr.push(arr);
          }
        }
      });
      tempArr = [...devArr];
    }
    if (tempApFilterKeys.Term != "All") {
      tempArr = tempArr.filter((arr) => {
        return arr.Term == tempApFilterKeys.Term;
      });
    }

    let lastIndex: number = 1 * totalPageItems;
    let firstIndex: number = lastIndex - totalPageItems;
    let paginatedItems = tempArr.slice(firstIndex, lastIndex);
    currentpage = 1;

    filterKeysAfterModified(tempArr);
    setApFilterOptions({ ...tempApFilterKeys });
    setApData(tempArr);
    setdisplayData(paginatedItems);
    setApCurrentPage(1);
  };
  const dateFormater = (date: Date): string => {
    return date ? moment(date).format("DD/MM/YYYY") : "";
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
  const onChangeHandler = (key: string, value: string) => {
    var testdataresponse = apResponseData;
    testdataresponse[key] = value;

    setApResponseData({ ...testdataresponse });
  };
  const apAddItem = () => {
    let product = [];

    if (apResponseData.product != null) {
      product = apMasterProducts.filter((prod) => {
        return prod.productName == apResponseData.product;
      });
    }

    const requestdata = {
      Title: apResponseData.projectOrTask ? apResponseData.projectOrTask : "",
      Status: "Scheduled",
      Master_x0020_ProjectId: product.length > 0 ? product[0].ProductId : null,
      ProjectOwnerId: apResponseData.manager ? apResponseData.manager : null,
      ProjectLeadId: apResponseData.developer
        ? { results: [apResponseData.developer] }
        : { results: [] },
      Year: apResponseData.year ? apResponseData.year : null,
      Term: apResponseData.term ? apResponseData.term : null,
      BusinessArea: apResponseData.businessArea
        ? apResponseData.businessArea
        : null,
      BA_x0020_acronyms: apResponseData.businessArea
        ? BAacronymsCollection.filter((BAacronym) => {
            return BAacronym.Name == apResponseData.businessArea;
          })[0].ShortName
        : null,
      ProjectType: apResponseData.typeOfProject
        ? apResponseData.typeOfProject
        : null,
      StartDate: apResponseData.startDate
        ? moment(apResponseData.startDate).format("YYYY-MM-DD")
        : null,
      PlannedEndDate: apResponseData.endDate
        ? moment(apResponseData.endDate).format("YYYY-MM-DD")
        : null,
    };

    sharepointWeb.lists
      .getByTitle(ListNameURL)
      .items.add(requestdata)
      .then((e) => {
        sharepointWeb.lists
          .getByTitle(ListNameURL)
          .items.getById(e.data.Id)
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
          .get()
          .then((item) => {
            let tempMasterArr = [...apMasterData];
            let newItemAddedtoArr = [];
            let arrAfterAddApData = allitemsArrayFormatter(
              item,
              newItemAddedtoArr
            );

            Array.prototype.push.apply(arrAfterAddApData, tempMasterArr);

            filterKeysAfterModified(arrAfterAddApData);

            let lastIndex: number = 1 * totalPageItems;
            let firstIndex: number = lastIndex - totalPageItems;
            let paginatedItems = arrAfterAddApData.slice(firstIndex, lastIndex);

            setApModalBoxVisibility({
              condition: false,
              action: "",
              selectedItem: [],
            });

            setApPopup("Success");
            setApMasterData([...arrAfterAddApData]);
            setApData([...arrAfterAddApData]);
            setdisplayData([...paginatedItems]);
            setApCurrentPage(1);
            setApShowMessage(apErrorStatus);
            setApResponseData(responseData);
            setApModelBoxDrpDwnToTxtBox(false);
            setApOnSubmitLoader(false);
            setTimeout(() => {
              setApPopup("Close");
            }, 2000);
          })
          .catch(apErrorFunction);
      })
      .catch(apErrorFunction);
  };
  const apUpdateItem = (id: number) => {
    let product = [];

    if (apResponseData.product != null) {
      product = apMasterProducts.filter((prod) => {
        return prod.productName == apResponseData.product;
      });
    }

    const requestdata = {
      Title: apResponseData.projectOrTask ? apResponseData.projectOrTask : "",
      Master_x0020_ProjectId: product.length > 0 ? product[0].ProductId : null,
      ProjectOwnerId: apResponseData.manager ? apResponseData.manager : null,
      ProjectLeadId: apResponseData.developer
        ? { results: [apResponseData.developer] }
        : { results: [] },
      Year: apResponseData.year ? apResponseData.year : null,
      Term: apResponseData.term ? apResponseData.term : null,
      BusinessArea: apResponseData.businessArea
        ? apResponseData.businessArea
        : null,
      BA_x0020_acronyms: apResponseData.businessArea
        ? BAacronymsCollection.filter((BAacronym) => {
            return BAacronym.Name == apResponseData.businessArea;
          })[0].ShortName
        : null,
      ProjectType: apResponseData.typeOfProject
        ? apResponseData.typeOfProject
        : null,
      StartDate: apResponseData.startDate
        ? moment(apResponseData.startDate).format("YYYY-MM-DD")
        : null,
      PlannedEndDate: apResponseData.endDate
        ? moment(apResponseData.endDate).format("YYYY-MM-DD")
        : null,
    };

    sharepointWeb.lists
      .getByTitle(ListNameURL)
      .items.getById(id)
      .update(requestdata)
      .then(() => {
        sharepointWeb.lists
          .getByTitle(ListNameURL)
          .items.getById(id)
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
          .get()
          .then((item) => {
            let tempMasterArr = [...apMasterData];
            let updatedItemtoArr = [];
            let targetIndex = tempMasterArr.findIndex((arr) => arr.ID == id);
            tempMasterArr.splice(targetIndex, 1);
            let arrAfterUpdateApData = allitemsArrayFormatter(
              item,
              updatedItemtoArr
            );
            Array.prototype.push.apply(arrAfterUpdateApData, tempMasterArr);

            filterKeysAfterModified(arrAfterUpdateApData);

            let lastIndex: number = 1 * totalPageItems;
            let firstIndex: number = lastIndex - totalPageItems;
            let paginatedItems = arrAfterUpdateApData.slice(
              firstIndex,
              lastIndex
            );

            setApModalBoxVisibility({
              condition: false,
              action: "",
              selectedItem: [],
            });

            setApPopup("Update");
            setApMasterData([...arrAfterUpdateApData]);
            setApData([...arrAfterUpdateApData]);
            setdisplayData([...paginatedItems]);
            setApCurrentPage(1);
            setApShowMessage(apErrorStatus);
            setApResponseData(responseData);
            setApOnSubmitLoader(false);
            setTimeout(() => {
              setApPopup("Close");
            }, 2000);
          })
          .catch(apErrorFunction);
      })
      .catch(apErrorFunction);
  };
  const apDeleteItem = (id: number) => {
    sharepointWeb.lists
      .getByTitle(ListNameURL)
      .items.getById(id)
      .delete()
      .then(() => {
        let tempMasterArr = [...apMasterData];
        let targetIndex = tempMasterArr.findIndex((arr) => arr.ID == id);
        tempMasterArr.splice(targetIndex, 1);

        let temp_ap_arr = [...apData];
        let targetIndexapdata = temp_ap_arr.findIndex((arr) => arr.ID == id);
        temp_ap_arr.splice(targetIndexapdata, 1);

        filterKeysAfterModified(temp_ap_arr);
        setApPopup("Delete");
        setApMasterData(tempMasterArr);
        setApData([...temp_ap_arr]);
        paginatewithdata(apcurrentPage, temp_ap_arr);
        setApOnDeleteLoader(false);
        setApDeletePopup({ condition: false, targetId: 0 });
        setTimeout(() => {
          setApPopup("Close");
        }, 2000);
      })
      .catch(apErrorFunction);
  };
  const apValidationFunction = () => {
    let isError = false;

    let errorStatus = {
      businessAreaError: "",
      projectOrTaskError: "",
    };

    if (!apResponseData.businessArea) {
      isError = true;
      errorStatus.businessAreaError = "Please select business area";
    }
    if (!apResponseData.projectOrTask) {
      isError = true;
      errorStatus.projectOrTaskError = "Please select project or task";
    }

    if (!isError) {
      if (apModalBoxVisibility.action == "Add") {
        setApOnSubmitLoader(true);
        apAddItem();
      } else if (apModalBoxVisibility.action == "Update") {
        setApOnSubmitLoader(true);
        apUpdateItem(apResponseData.ID);
      }
    } else {
      setApShowMessage(errorStatus);
    }
  };
  const paginate = (pagenumber) => {
    let lastIndex: number = pagenumber * totalPageItems;
    let firstIndex: number = lastIndex - totalPageItems;
    let paginatedItems = apData.slice(firstIndex, lastIndex);
    currentpage = pagenumber;
    setdisplayData(paginatedItems);
    setApCurrentPage(pagenumber);
  };
  const paginatewithdata = (pagenumber, data) => {
    let lastIndex: number = pagenumber * totalPageItems;
    let firstIndex: number = lastIndex - totalPageItems;
    let paginatedItems = data.slice(firstIndex, lastIndex);
    currentpage = pagenumber;
    if (paginatedItems.length > 0) {
      setdisplayData(paginatedItems);
      setApCurrentPage(pagenumber);
    } else {
      paginate(pagenumber - 1);
    }
  };
  const generateExcel = () => {
    let arrExport = apMasterData;
    const workbook = new Excel.Workbook();
    const worksheet = workbook.addWorksheet("My Sheet");
    worksheet.columns = [
      // { header: "ID", key: "id", width: 25 },
      { header: "Business area", key: "businessArea", width: 25 },
      { header: "Term", key: "term", width: 25 },
      { header: "Hours", key: "hours", width: 25 },
      { header: "Start date", key: "startDate", width: 25 },
      { header: "End date", key: "endDate", width: 25 },
      { header: "Product or solution", key: "product", width: 60 },
      { header: "Type of deliverbale", key: "typeOfDeliverable", width: 20 },
      { header: "Project or task", key: "projectOrTask", width: 40 },
      { header: "Status", key: "status", width: 30 },
      { header: "Manager", key: "manager", width: 30 },
      { header: "Developer", key: "developer", width: 30 },
    ];
    arrExport.forEach((item) => {
      worksheet.addRow({
        // id: item.ID,
        businessArea: item.BusinessArea ? item.BusinessArea : "",
        term: item.Term ? parseInt(item.Term) : "",
        hours: item.Hours ? item.Hours : "",
        startDate: item.StartDate ? item.StartDate : "",
        endDate: item.EndDate ? item.EndDate : "",
        product: item.Product ? item.Product : "",
        typeOfDeliverable: item.TypeOfProject ? item.TypeOfProject : "",
        projectOrTask: item.ProjectOrTask ? item.ProjectOrTask : "",
        status: item.StatusStage ? item.StatusStage : "",
        manager: item.PMName ? item.PMName.name : "",
        developer: item.DNames.length > 0 ? item.DNames[0].name : "",
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
      // "L1",
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
      // "L1",
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
          `AnnualPlan-${new Date().toLocaleString()}.xlsx`
        )
      )
      .catch((err) => console.log("Error writing excel export", err));
  };
  const AddSuccessPopup = () => (
    <MessageBar messageBarType={MessageBarType.success} isMultiline={false}>
      Annual plan has been successfully added !!!
    </MessageBar>
  );
  const UpdateSuccessPopup = () => (
    <MessageBar messageBarType={MessageBarType.success} isMultiline={false}>
      Annual plan has been successfully updated !!!
    </MessageBar>
  );
  const DeleteSuccessPopup = () => (
    <MessageBar messageBarType={MessageBarType.warning} isMultiline={false}>
      Annual plan has been successfully deleted !!!
    </MessageBar>
  );
  const ErrorPopup = () => (
    <MessageBar messageBarType={MessageBarType.error} isMultiline={false}>
      Something when error, please contact system admin.
    </MessageBar>
  );

  useEffect(() => {
    getAllUsers();
    getApData();
    getAllOptions();
  }, [apReRender]);

  return (
    <div style={{ padding: "5px 15px" }}>
      {apStartUpLoader ? <CustomLoader /> : null}
      <div className={styles.apHeaderSection}>
        <div>
          {apPopup == "Success"
            ? AddSuccessPopup()
            : apPopup == "Update"
            ? UpdateSuccessPopup()
            : apPopup == "Delete"
            ? DeleteSuccessPopup()
            : apPopup == "Error"
            ? ErrorPopup()
            : ""}
        </div>
        <div className={styles.apHeader}>Annual plan</div>
        <div style={{ display: "flex", justifyContent: "space-between" }}>
          <div className={styles.apAddBtn}>
            <a
              onClick={(_) => {
                setApData(apMasterData);
                filterKeysAfterModified(apMasterData);
                setApFilterOptions({ ...apFilterKeys });
                paginatewithdata(1, apMasterData);
                setApModalBoxVisibility({
                  condition: true,
                  action: "Add",
                  selectedItem: [],
                });
              }}
            >
              Add deliverable
            </a>
          </div>
          <div style={{ display: "flex" }}>
            <Label styles={apLabelStyles} style={{ paddingTop: 13 }}>
              Number of records :{" "}
              <b style={{ color: "#038387" }}>{apMasterData.length}</b>
            </Label>
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
              }}
            >
              <Icon
                style={{
                  color: "#1D6F42",
                }}
                iconName="ExcelDocument"
                className={apIconStyleClass.export}
              />
              Export as XLS
            </Label>
          </div>
        </div>
        {/* Dropdown Section */}
        <div style={{ display: "flex" }}>
          <div>
            <Label styles={apLabelStyles}>Search</Label>
            <SearchBox
              placeholder="Find Project or Task"
              styles={
                apFilterOptions.ProjectOrTaskSearch == ""
                  ? apSearchBoxStyles
                  : apActiveSearchBoxStyles
              }
              value={apFilterOptions.ProjectOrTaskSearch}
              onChange={(e, value) => {
                listFilter("ProjectOrTaskSearch", value);
              }}
            />
          </div>
          <div>
            <Label styles={apLabelStyles}>Business area</Label>
            <Dropdown
              placeholder="Select an option"
              options={apDropDownOptions.baOptns}
              styles={
                apFilterOptions.BusinessArea == "All"
                  ? apDropdownStyles
                  : apActiveDropdownStyles
              }
              onChange={(e, option: any) => {
                listFilter("BusinessArea", option["key"]);
              }}
              selectedKey={apFilterOptions.BusinessArea}
            />
          </div>
          <div>
            <Label styles={apLabelStyles}>Type of deliverable</Label>
            <Dropdown
              selectedKey={apFilterOptions.TypeOfProject}
              placeholder="Select an option"
              options={apDropDownOptions.todOptns}
              styles={
                apFilterOptions.TypeOfProject == "All"
                  ? apDropdownStyles
                  : apActiveDropdownStyles
              }
              onChange={(e, option: any) => {
                listFilter("TypeOfProject", option["key"]);
              }}
            />
          </div>
          <div>
            <Label styles={apLabelStyles}>Project or task</Label>
            <Dropdown
              selectedKey={apFilterOptions.ProjectOrTask}
              placeholder="Select an option"
              options={apDropDownOptions.potOptns}
              dropdownWidth={"auto"}
              styles={
                apFilterOptions.ProjectOrTask == "All"
                  ? apDropdownStyles
                  : apActiveDropdownStyles
              }
              onChange={(e, option: any) => {
                listFilter("ProjectOrTask", option["key"]);
              }}
            />
          </div>
          <div>
            <Label styles={apLabelStyles}>Manager</Label>
            <Dropdown
              selectedKey={apFilterOptions.PM}
              placeholder="Select an option"
              options={apDropDownOptions.managerOptns}
              styles={
                apFilterOptions.PM == "All"
                  ? apDropdownStyles
                  : apActiveDropdownStyles
              }
              onChange={(e, option: any) => {
                listFilter("PM", option["key"]);
              }}
            />
          </div>
          <div>
            <Label styles={apLabelStyles}>Developer</Label>
            <Dropdown
              selectedKey={apFilterOptions.D}
              placeholder="Select an option"
              options={apDropDownOptions.developerOptns}
              styles={
                apFilterOptions.D == "All"
                  ? apDropdownStyles
                  : apActiveDropdownStyles
              }
              onChange={(e, option: any) => {
                listFilter("D", option["key"]);
              }}
            />
          </div>
          <div>
            <Label styles={apLabelStyles}>Term</Label>
            <Dropdown
              selectedKey={apFilterOptions.Term}
              multiSelect={false}
              placeholder="Select an option"
              options={apDropDownOptions.termOptns}
              styles={
                apFilterOptions.Term == "All"
                  ? apDropdownStyles
                  : apActiveDropdownStyles
              }
              onChange={(e, option: any) => {
                listFilter("Term", option["key"]);
              }}
            />
          </div>
          <div>
            <div>
              <Icon
                iconName="Refresh"
                className={apIconStyleClass.refresh}
                onClick={() => {
                  setApData(apMasterData);
                  filterKeysAfterModified(apMasterData);
                  setApFilterOptions({ ...apFilterKeys });
                  paginatewithdata(1, apMasterData);
                }}
              />
            </div>
          </div>
        </div>
        {/* Dropdown Section */}
      </div>

      <div>
        <DetailsList
          items={displayData}
          columns={apColumns}
          setKey="set"
          layoutMode={DetailsListLayoutMode.justified}
          selectionMode={SelectionMode.none}
          styles={gridStyles}
        />
      </div>
      <div>
        {/* {apStartUpLoader ? (
          // <Spinner
          //   size={SpinnerSize.large}
          //   label="Please wait..."
          //   labelPosition="left"
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
        ) :  */}
        {displayData.length > 0 ? (
          <div
            style={{
              display: "flex",
              justifyContent: "center",
              margin: "10px 0",
            }}
          >
            <Pagination
              currentPage={apcurrentPage}
              totalPages={
                apData.length > 0
                  ? Math.ceil(apData.length / totalPageItems)
                  : 1
              }
              onChange={(page) => {
                paginate(page);
              }}
            />
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
      </div>
      <div>
        {apModalBoxVisibility.condition ? (
          <Modal isOpen={apModalBoxVisibility.condition} isBlocking={true}>
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
                {apModalBoxVisibility.action == "Add"
                  ? "New deliverable "
                  : apModalBoxVisibility.action == "Update"
                  ? "Update deliverable"
                  : ""}
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
                    label="Business area"
                    required={true}
                    errorMessage={apShowMessage.businessAreaError}
                    selectedKey={apResponseData.businessArea}
                    placeholder="Select an option"
                    options={apModalBoxDropDownOptions.baOptns}
                    styles={apModalBoxDrpDwnCalloutStyles}
                    onChange={(e, option: any) => {
                      onChangeHandler("businessArea", option["key"]);
                    }}
                  />
                </div>
                <div>
                  <Dropdown
                    label="Type of deliverable"
                    selectedKey={apResponseData.typeOfProject}
                    placeholder="Select an option"
                    options={apModalBoxDropDownOptions.todOptns}
                    styles={apModalBoxDrpDwnCalloutStyles}
                    onChange={(e, option: any) => {
                      onChangeHandler("typeOfProject", option["key"]);
                    }}
                  />
                </div>
                <div>
                  <Dropdown
                    label="Term"
                    selectedKey={apResponseData.term}
                    placeholder="Select an option"
                    options={apModalBoxDropDownOptions.termOptns}
                    styles={apModalBoxDropdownStyles}
                    onChange={(e, option: any) => {
                      onChangeHandler("term", option["key"]);
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
                  <Dropdown
                    label="Product or solution"
                    selectedKey={apResponseData.product}
                    placeholder="Select an option"
                    options={apModalBoxDropDownOptions.productOptns}
                    styles={apModalBoxDrpDwnCalloutStyles}
                    onChange={(e, option: any) => {
                      onChangeHandler("product", option["key"]);
                    }}
                  />
                </div>
                <div>
                  <DatePicker
                    label="Start date"
                    placeholder="Select a date..."
                    ariaLabel="Select a date"
                    formatDate={dateFormater}
                    value={apResponseData.startDate}
                    styles={apModalBoxDatePickerStyles}
                    onSelectDate={(value: any) => {
                      onChangeHandler("startDate", value);
                    }}
                  />
                </div>
                <div>
                  <DatePicker
                    label="End date"
                    placeholder="Select a date..."
                    ariaLabel="Select a date"
                    styles={apModalBoxDatePickerStyles}
                    formatDate={dateFormater}
                    value={apResponseData.endDate}
                    onSelectDate={(value: any) => {
                      onChangeHandler("endDate", value);
                    }}
                  />
                </div>
              </div>
              {apModalBoxVisibility.action == "Add" ? (
                <div
                  style={{
                    display: "flex",
                    alignItems: "flex-start",
                    justifyContent: "space-between",
                  }}
                >
                  <div>
                    {apModelBoxDrpDwnToTxtBox ? (
                      <TextField
                        label="Project or task"
                        placeholder="Add project or task"
                        errorMessage={apShowMessage.projectOrTaskError}
                        required={true}
                        defaultValue={apResponseData.projectOrTask}
                        styles={apTxtBoxStyles}
                        onChange={(e, value: string) => {
                          onChangeHandler("projectOrTask", value);
                        }}
                      />
                    ) : (
                      <Dropdown
                        label="Project or task"
                        selectedKey={apResponseData.projectOrTask}
                        placeholder="Select project or task"
                        errorMessage={apShowMessage.projectOrTaskError}
                        required={true}
                        options={apModalBoxDropDownOptions.potOptns}
                        styles={apModalBoxDrpDwnCalloutStyles}
                        style={{ width: "850px" }}
                        onChange={(e, option: any) => {
                          onChangeHandler("projectOrTask", option["key"]);
                        }}
                      />
                    )}
                  </div>
                  <div>
                    {apModelBoxDrpDwnToTxtBox ? (
                      <Checkbox
                        label="New"
                        styles={apModalBoxCheckBoxStyles}
                        checked={apModelBoxDrpDwnToTxtBox}
                        onChange={(e) => {
                          onChangeHandler("projectOrTask", "");
                          setApModelBoxDrpDwnToTxtBox(
                            !apModelBoxDrpDwnToTxtBox
                          );
                        }}
                      />
                    ) : (
                      <Checkbox
                        label="New"
                        styles={apModalBoxCheckBoxStyles}
                        checked={apModelBoxDrpDwnToTxtBox}
                        onChange={(e) => {
                          onChangeHandler("projectOrTask", "");
                          setApModelBoxDrpDwnToTxtBox(
                            !apModelBoxDrpDwnToTxtBox
                          );
                        }}
                      />
                    )}
                  </div>
                </div>
              ) : apModalBoxVisibility.action == "Update" ? (
                <div style={{ display: "flex" }}>
                  <div>
                    <TextField
                      label="Project or task"
                      placeholder="Add project or task"
                      errorMessage={apShowMessage.projectOrTaskError}
                      defaultValue={apResponseData.projectOrTask}
                      required={true}
                      styles={apTxtBoxStyles}
                      onChange={(e, value: string) => {
                        onChangeHandler("projectOrTask", value);
                      }}
                    />
                  </div>
                </div>
              ) : (
                ""
              )}
              <div
                style={{
                  display: "flex",
                  alignItems: "flex-start",
                  justifyContent: "flex-start",
                }}
              >
                <div>
                  <Dropdown
                    label="Year"
                    selectedKey={apResponseData.year}
                    placeholder="Select an option"
                    options={apModalBoxDropDownOptions.yearOptns}
                    styles={apModalBoxDrpDwnCalloutStyles}
                    onChange={(e, option: any) => {
                      onChangeHandler("year", option["key"]);
                    }}
                  />
                </div>
                <div>
                  <Label
                    style={{
                      transform: "translate(20px, 10px)",
                      // marginTop: "3px",
                    }}
                  >
                    Manager
                  </Label>
                  <NormalPeoplePicker
                    className={apModalBoxPP}
                    onResolveSuggestions={GetUserDetails}
                    itemLimit={1}
                    defaultSelectedItems={peopleList.filter((people) => {
                      return people.ID == apResponseData.manager;
                    })}
                    onChange={(selectedUser) => {
                      selectedUser.length != 0
                        ? onChangeHandler("manager", selectedUser[0]["ID"])
                        : onChangeHandler("manager", "");
                    }}
                  />
                </div>
                <div>
                  <Label
                    style={{
                      transform: "translate(20px, 10px)",
                    }}
                  >
                    Developer
                  </Label>
                  <NormalPeoplePicker
                    className={apModalBoxPP}
                    onResolveSuggestions={GetUserDetails}
                    itemLimit={1}
                    defaultSelectedItems={peopleList.filter((people) => {
                      return people.ID == apResponseData.developer;
                    })}
                    onChange={(selectedUser) => {
                      selectedUser.length != 0
                        ? onChangeHandler("developer", selectedUser[0]["ID"])
                        : onChangeHandler("developer", "");
                    }}
                  />
                </div>
              </div>
              <div className={styles.apModalBoxButtonSection}>
                <button
                  className={styles.apModalBoxSubmitBtn}
                  onClick={(_) => {
                    apValidationFunction();
                  }}
                  style={{ display: "flex" }}
                >
                  {apOnSubmitLoader ? (
                    <Spinner />
                  ) : apModalBoxVisibility.action == "Add" ? (
                    <span>
                      <Icon
                        iconName="Save"
                        // style={{ marginRight: 12, marginLeft: -12 }}
                        style={{ position: "relative", top: 3, left: -8 }}
                      />
                      {"Submit"}
                    </span>
                  ) : apModalBoxVisibility.action == "Update" ? (
                    <span>
                      <Icon
                        iconName="Save"
                        // style={{ marginRight: 12, marginLeft: -12 }}
                        style={{ position: "relative", top: 3, left: -8 }}
                      />
                      {"Update"}
                    </span>
                  ) : (
                    ""
                  )}
                </button>
                <button
                  className={styles.apModalBoxBackBtn}
                  onClick={(_) => {
                    setApResponseData(responseData);
                    setApShowMessage(apErrorStatus);
                    setApModelBoxDrpDwnToTxtBox(false);
                    setApModalBoxVisibility({
                      condition: false,
                      action: "",
                      selectedItem: [],
                    });
                  }}
                >
                  <span>
                    {" "}
                    <Icon
                      iconName="Cancel"
                      style={{ position: "relative", top: 3, left: -8 }}
                      // style={{ marginRight: 12, marginLeft: -12 }}
                    />
                    Close
                  </span>
                </button>
              </div>
            </div>
          </Modal>
        ) : (
          ""
        )}
      </div>
      {/* Delete Popup */}
      <div>
        {apDeletePopup.condition ? (
          <Modal isOpen={apDeletePopup.condition} isBlocking={true}>
            <div
              style={{
                display: "flex",
                justifyContent: "center",
                alignItems: "center",
                marginTop: "30px",
                width: "450px",
              }}
            >
              <div
                style={{
                  display: "flex",
                  alignItems: "center",
                  justifyContent: "flex-start",
                  flexDirection: "column",
                  marginBottom: "10px",
                }}
              >
                <Label className={styles.deletePopupTitle}>
                  Delete Project or task
                </Label>
                <Label className={styles.deletePopupDesc}>
                  Are you sure you want to delete this project or task?
                </Label>
              </div>
            </div>
            <div className={styles.apDeletePopupBtnSection}>
              <button
                onClick={(_) => {
                  setApOnDeleteLoader(true);
                  apDeleteItem(apDeletePopup.targetId);
                }}
                className={styles.apDeletePopupYesBtn}
              >
                {apOnDeleteLoader ? (
                  <Spinner />
                ) : (
                  //<CustomLoader />
                  "Yes"
                )}
              </button>
              <button
                onClick={(_) => {
                  setApDeletePopup({ condition: false, targetId: 0 });
                }}
                className={styles.apDeletePopupNoBtn}
              >
                No
              </button>
            </div>
          </Modal>
        ) : (
          ""
        )}
      </div>
    </div>
  );
};

export default AnnualPlan;
