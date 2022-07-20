import * as React from "react";
import { useState, useEffect } from "react";
import * as moment from "moment";
import { Web } from "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/fields";
import "@pnp/sp/webs";
import "@pnp/sp/site-users/web";

import {
  DetailsList,
  DetailsListLayoutMode,
  IDetailsListStyles,
  SelectionMode,
  Icon,
  Label,
  ILabelStyles,
  Toggle,
  IToggleStyles,
  SearchBox,
  ISearchBoxStyles,
  Dropdown,
  IDropdownStyles,
  NormalPeoplePicker,
  IBasePickerStyles,
  Persona,
  PersonaPresence,
  PersonaSize,
  TextField,
  ITextFieldStyles,
  TooltipHost,
  TooltipOverflowMode,
  Rating,
  RatingSize,
  IRatingStyles,
  Modal,
  IModalStyles,
  Spinner,
  MessageBar,
  MessageBarType,
} from "@fluentui/react";
import ReactQuill from "react-quill";
import "react-quill/dist/quill.snow.css";
import "../ExternalRef/styleSheets/ProdStyles.css";
import { mergeStyles, mergeStyleSets } from "@fluentui/react/lib/Styling";
import styles from "./InnovationHubIntranet.module.scss";
import CustomLoader from "./CustomLoader";

const DocumentReview = (props: any) => {
  // Variable-Declaration-Section Starts
  const sharepointWeb = Web(props.URL);
  const drListName = "Review Log";

  const drAllitems = [];
  const drAllpeoples = [];
  const drColumns = [
    {
      key: "Request",
      name: "Request",
      fieldName: "Request",
      minWidth: 60,
      maxWidth: 60,
    },
    {
      key: "FileName",
      name: "FileName",
      fieldName: "FileName",
      minWidth: 240,
      maxWidth: 240,
      onRender: (item) => (
        <div style={{ cursor: "pointer" }}>
          <TooltipHost
            id={item.ID}
            content={item.FileName}
            overflowMode={TooltipOverflowMode.Parent}
          >
            <span aria-describedby={item.ID}>{item.FileName}</span>
          </TooltipHost>
        </div>
      ),
    },
    {
      key: "Sent",
      name: "Sent",
      fieldName: "Sent",
      minWidth: 80,
      maxWidth: 80,
      onRender: (item) => moment(item.Sent).format("DD/MM/YYYY"),
    },
    {
      key: "Response",
      name: "Response",
      fieldName: "Response",
      minWidth: 80,
      maxWidth: 80,
    },
    {
      key: "User",
      name: "User",
      fieldName: "User",
      minWidth: 40,
      maxWidth: 40,

      onRender: (item) => (
        <>
          <div
            title={item.UserDetails.UserName}
            style={{
              marginTop: "-6px",
            }}
          >
            <Persona
              size={PersonaSize.size32}
              presence={PersonaPresence.none}
              imageUrl={
                "/_layouts/15/userphoto.aspx?size=S&username=" +
                `${item.UserDetails.UserEmail}`
              }
            />
          </div>
        </>
      ),
    },
    {
      key: "Actions",
      name: "Actions",
      fieldName: "Actions",
      minWidth: 50,
      maxWidth: 50,

      onRender: (item) => (
        <>
          <Icon
            iconName="ChevronRightMed"
            className={drIconStyleClass.DetailArrow}
            onClick={async () => {
              let targetObj = drData.filter((data) => {
                return data.ID == item.ID;
              });
              await setDrReviewFormDisplay({
                condition: false,
                selectedItem: {},
              });
              await setDrReviewFormOptionDisplay({
                condition: false,
                selectedOption: null,
                issuesCategory: {
                  issues: "",
                  issuesSeverity: "",
                  issueRepeated: false,
                },
                rating: 0,
              });
              await setDrReviewFormDisplay({
                condition: true,
                selectedItem: targetObj[0],
              });
            }}
          />
        </>
      ),
    },
  ];
  const drDrpDwnOptns = {
    viewOptns: [
      { key: "All", text: "All" },
      { key: "Pending", text: "Pending" },
      { key: "Pending edit", text: "Pending edit" },
      { key: "Send by me", text: "Send by me" },
      { key: "Responded by me", text: "Responded by me" },
      { key: "Last 30 days", text: "Last 30 days" },
    ],
    toOptns: [
      { key: "Me", text: "Me" },
      { key: "Me or Me Cc'd", text: "Me or Me Cc'd" },
      { key: "Anyone", text: "Anyone" },
    ],
    requestOptns: [{ key: "All", text: "All" }],
    responseOptns: [{ key: "All", text: "All" }],
  };
  const filters = {
    view: "Pending",
    to: "Me",
    request: "All",
    response: "All",
    toUser: "",
    fromUser: "",
    fileName: "",
    product: "",
  };
  const modules = {
    toolbar: [
      [
        {
          header: [1, 2, 3, false],
        },
      ],
      ["bold", "italic", "underline"],
      [
        {
          color: [],
        },
        {
          background: [],
        },
      ],
      [
        {
          list: "ordered",
        },
        {
          list: "bullet",
        },
        {
          indent: "-1",
        },
        {
          indent: "+1",
        },
      ],
      ["clean"],
    ],
  };
  const formats = [
    "header",
    "bold",
    "italic",
    "underline",
    "list",
    "bullet",
    "indent",
    "background",
    "color",
  ];
  // Variable-Declaration-Section Ends
  // Styles-Section Starts
  const drLabelStyles: Partial<ILabelStyles> = {
    root: {
      width: 150,
      marginRight: 10,
      fontSize: "13px",
      color: "#323130",
    },
  };
  const drToggleStyles: Partial<IToggleStyles> = {
    root: {
      marginTop: "15px",
    },
    pill: {
      border: "none",
      backgroundColor: "#2392B2",
      hover: {
        backgroundColor: "#2392B2",
      },
    },
    thumb: {
      border: "none",
      backgroundColor: "#FFF",
      hover: {
        backgroundColor: "#FFF",
      },
    },
  };
  const drDetailsListStyles: Partial<IDetailsListStyles> = {
    root: {
      width: 670,
      selectors: {
        ".ms-DetailsRow-cell": {
          height: 40,
        },
      },
    },
    contentWrapper: {
      height: 430,
      overflowX: "hidden",
      overflowY: "scroll",
    },
  };
  const drDropdownStyles: Partial<IDropdownStyles> = {
    root: {
      width: 165,
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
  };
  const drReviewFormDropDownStyles: Partial<IDropdownStyles> = {
    root: {
      width: 240,
      margin: "10px 10px 10px 0",
    },
    title: {
      height: "36px",
      padding: "1px 10px",
    },
    caretDown: {
      fontSize: 14,
      padding: "3px",
      color: "#000",
      fontWeight: "bold",
    },
  };
  const drSearchBoxStyles: Partial<ISearchBoxStyles> = {
    root: {
      width: 165,
      marginRight: "15px",
      backgroundColor: "#F5F5F7",
      outline: "none",
      border: "1px solid #E8E8EA",
      borderRadius: "4px",
    },
    icon: { fontSize: 14, color: "#000" },
  };
  const drModalStyles: Partial<IModalStyles> = {
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
  const drModalTextFields: Partial<ITextFieldStyles> = {
    root: { width: "93%", margin: "10px 20px" },
    fieldGroup: {
      height: 40,
    },
  };
  const drReviewFormPP: Partial<IBasePickerStyles> = {
    root: { width: "600px", margin: "10px 0px" },
    input: {
      height: 36,
      padding: "0px 10px !important",
    },
    itemsWrapper: {
      padding: "0px 5px !important",
    },
  };
  const drModalBoxPP: Partial<IBasePickerStyles> = {
    root: {
      width: "93%",
      margin: "10px 20px",
    },
    itemsWrapper: {
      height: "30px !important",
      width: "100% !important",
      padding: "0px 3px !important",
    },
    text: {
      height: "40px !important",
      padding: "4px 3px !important",
      // paddingTop: 4,
      // paddingBottom: 4,
      width: "100% !important",
    },
  };
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
  const drIconStyleClass = mergeStyleSets({
    DetailArrow: [
      {
        fontSize: 25,
        height: 14,
        width: 17,
        color: "#038387",
        margin: "0 7px",
        cursor: "pointer",
      },
    ],
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
  });
  // Styles-Section Ends
  // States-Declaration Starts
  const [drReRender, setDrReRender] = useState(true);
  const [currentUser, setCurrentUser] = useState({});
  const [drMasterData, setDrMasterData] = useState(drAllitems);
  const [drData, setDrData] = useState(drAllitems);
  const [peopleList, setPeopleList] = useState(drAllpeoples);
  const [documentReviewAdmins, setDocumentReviewAdmins] = useState([]);
  const [drDropDownOptions, setDrDropDownOptions] = useState(drDrpDwnOptns);
  const [drFilters, setDrFilters] = useState(filters);
  const [drReviewFormDisplay, setDrReviewFormDisplay] = useState({
    condition: false,
    selectedItem: {},
  });
  const [drReviewFormOptionDisplay, setDrReviewFormOptionDisplay] = useState({
    condition: false,
    selectedOption: null,
    issuesCategory: { issues: "", issuesSeverity: "", issueRepeated: false },
    rating: 4,
  });
  const [drReallocatePopup, setDrReallocatePopup] = useState({
    condition: false,
    allocatedUser: null,
  });
  const [drReallocateUser, setDrReallocateUser] = useState({});
  const [drCancelRequestPopup, setDrCancelRequestPopup] = useState(false);
  const [drCancelReason, setDrCancelReason] = useState("");
  const [drSignOffPopup, setDrSignOffPopup] = useState(false);
  const [drSignOffOptions, setDrSignOffOptions] = useState({
    assignTo: null,
    signOffComments: "",
    publishRequestComments: "",
  });
  const [drLoader, setDrLoader] = useState("noLoader");
  const [drPopup, setDrPopup] = useState("close");
  // States-Declaration Ends
  //Function-Section Starts
  const getAllDRAdmins = () => {
    sharepointWeb.siteGroups
      .getByName("Document Review Admins")
      .users.get()
      .then((users) => {
        let DRAdmins = [];
        users.forEach((user) => {
          DRAdmins.push({
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
        setDocumentReviewAdmins(DRAdmins);
      })
      .catch(drErrorFunction);
  };
  const getAllUsers = () => {
    sharepointWeb
      .siteUsers()
      .then((_allUsers) => {
        _allUsers.forEach((user) => {
          drAllpeoples.push({
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
        setPeopleList(drAllpeoples);
        drGetData(drAllpeoples);
      })
      .catch(drErrorFunction);
  };
  const drGetCurrentUserDetails = () => {
    sharepointWeb.currentUser
      .get()
      .then((user) => {
        let drCurrentUser = {
          Name: user.Title,
          Email: user.Email,
          Id: user.Id,
        };
        setCurrentUser(drCurrentUser);
      })
      .catch(drErrorFunction);
  };
  const drGetData = (peoples: any) => {
    sharepointWeb.lists
      .getByTitle(drListName)
      .items.select(
        "*",
        "FromUser/Title",
        "FromUser/Id",
        "FromUser/EMail",
        "ToUser/Title",
        "ToUser/Id",
        "ToUser/EMail",
        "CcEmail/Title",
        "CcEmail/Id",
        "CcEmail/EMail"
      )
      .expand("FromUser", "CcEmail", "ToUser")
      .top(5000)
      .orderBy("Modified", false)
      // .orderBy("ID", false)
      .get()
      .then((items) => {
        console.log(items);
        items.forEach((item) => {
          let tempCcEmails = [];
          if (item.CcEmailId) {
            peoples.forEach((people) => {
              item.CcEmail.forEach((email) => {
                if (people.ID == email.Id) {
                  tempCcEmails.push(people);
                }
              });
            });
          }

          drAllitems.push({
            ID: item.Id ? item.Id : "",
            Link: item.auditLink ? item.auditLink : "",
            Request: item.auditRequestType ? item.auditRequestType : "",
            FileName: item.Title ? item.Title : "",
            Sent: item.auditSent,
            Response: item.auditResponseType ? item.auditResponseType : "",
            UserDetails: {
              UserName: item.FromUser ? item.FromUser.Title : "",
              UserEmail: item.FromUser ? item.FromUser.EMail : "",
              UserId: item.FromUser ? item.FromUser.Id : "",
            },
            User: item.FromUser ? item.FromUser.Title : "",
            ToUserDetails: {
              UserName: item.ToUser ? item.ToUser.Title : "",
              UserEmail: item.ToUser ? item.ToUser.EMail : "",
              UserId: item.ToUser ? item.ToUser.Id : "",
            },
            ToUser: item.ToUser ? item.ToUser.Title : "",
            RequestComments: item.auditComments ? item.auditComments : "",
            ResponseComments: item.Response_x0020_Comments
              ? item.Response_x0020_Comments
              : "",
            CcEmailIds: item.CcEmailId ? item.CcEmailId : [],
            CcEmails: item.CcEmailId ? tempCcEmails : [],
            Product: item.ProductName ? item.ProductName : "",
            RepeatedIssue: item.FeedbackRepeated
              ? item.FeedbackRepeated
              : false,
            Rating: item.Rating ? item.Rating : 0,
            Created: item.Created,
            Modified: item.Modified,
          });
        });

        let drAllitemsAfterInitialFilter = drAllitems.filter((item) => {
          return (
            item.Response == "Pending" &&
            item.UserDetails.UserName ==
              props.spcontext.pageContext.user.displayName
          );
        });

        console.log(drAllitemsAfterInitialFilter);

        setDrMasterData([...drAllitems]);
        setDrData([...drAllitemsAfterInitialFilter]);
        setDrLoader("noLoader");
      })
      .catch(drErrorFunction);
  };
  const drGetAllOptions = () => {
    const _sortFilterKeys = (a, b) => {
      if (a.text < b.text) {
        return -1;
      }
      if (a.text > b.text) {
        return 1;
      }
      return 0;
    };

    //Request Choices
    sharepointWeb.lists
      .getByTitle(drListName)
      .fields.getByInternalNameOrTitle("auditRequestType")()
      .then((response) => {
        response["Choices"].forEach((choice) => {
          if (choice != null) {
            if (
              drDrpDwnOptns.requestOptns.findIndex((requestOptn) => {
                return requestOptn.key == choice;
              }) == -1
            ) {
              drDrpDwnOptns.requestOptns.push({
                key: choice,
                text: choice,
              });
            }
          }
        });
      })
      .then(() => {
        drDrpDwnOptns.requestOptns.shift();
        drDrpDwnOptns.requestOptns.sort(_sortFilterKeys);
        drDrpDwnOptns.requestOptns.unshift({ key: "All", text: "All" });
      })
      .catch(drErrorFunction);

    //Response Choices
    sharepointWeb.lists
      .getByTitle(drListName)
      .fields.getByInternalNameOrTitle("auditResponseType")()
      .then((response) => {
        response["Choices"].forEach((choice) => {
          if (choice != null) {
            if (
              drDrpDwnOptns.responseOptns.findIndex((responseOptn) => {
                return responseOptn.key == choice;
              }) == -1
            ) {
              drDrpDwnOptns.responseOptns.push({
                key: choice,
                text: choice,
              });
            }
          }
        });
      })
      .then(() => {
        drDrpDwnOptns.responseOptns.shift();
        drDrpDwnOptns.responseOptns.sort(_sortFilterKeys);
        drDrpDwnOptns.responseOptns.unshift({ key: "All", text: "All" });
      })
      .then(() => {
        setDrDropDownOptions(drDrpDwnOptns);
      })
      .catch(drErrorFunction);
  };
  const drhandleFilters = async (key: string, option: string) => {
    let _filters = { ...drFilters };
    _filters[key] = option;

    await filterItems(_filters);
    await setDrReviewFormDisplay({ condition: false, selectedItem: {} });

    // await queryGenerator(_filters);
  };
  const filterItems = async (filterKeys: any) => {
    let dataToBeFiltered = [...drMasterData];

    if (filterKeys.view && filterKeys.view != "All") {
      if (filterKeys.view == "Pending") {
        dataToBeFiltered = dataToBeFiltered.filter((arr) => {
          return arr.Response == "Pending";
        });
      }
      if (filterKeys.view == "Pending edit") {
        dataToBeFiltered = dataToBeFiltered.filter((arr) => {
          return (
            arr.Response == "Pending" &&
            (arr.Request == "Initial Edit" || arr.Request == "Final Edit")
          );
        });
      }
      if (filterKeys.view == "Send by me") {
        dataToBeFiltered = dataToBeFiltered.filter((arr) => {
          return arr.UserDetails.UserName == currentUser["Name"];
        });
      }
      if (filterKeys.view == "Responded by me") {
        dataToBeFiltered = dataToBeFiltered.filter((arr) => {
          return (
            arr.ToUserDetails.UserName == currentUser["Name"] &&
            arr.Response != "Pending"
          );
        });
      }
      if (filterKeys.view == "Last 30 days") {
        let todayDate = moment().format("YYYY-MM-DD");
        let last30Days = moment().subtract(30, "days").format("YYYY-MM-DD");
        dataToBeFiltered = dataToBeFiltered.filter((arr) => {
          return (
            moment(arr.Sent).format("YYYY-MM-DD") >= last30Days &&
            moment(arr.Sent).format("YYYY-MM-DD") <= todayDate
          );
        });
      }
    }

    if (filterKeys.to && filterKeys.to != "Anyone") {
      if (filterKeys.to == "Me") {
        dataToBeFiltered = dataToBeFiltered.filter((arr) => {
          return arr.UserDetails.UserName == currentUser["Name"];
        });
      }
      if (filterKeys.to == "Me or Me Cc'd") {
        dataToBeFiltered = dataToBeFiltered.filter((arr) => {
          return (
            arr.UserDetails.UserName == currentUser["Name"] ||
            arr.CcEmails.some((people) => {
              people.title == currentUser["Name"];
            }) == true
          );
        });
      }
    }

    if (filterKeys.request && filterKeys.request != "All") {
      dataToBeFiltered = dataToBeFiltered.filter((arr) => {
        return arr.Request == filterKeys.request;
      });
    }

    if (filterKeys.response && filterKeys.response != "All") {
      dataToBeFiltered = dataToBeFiltered.filter((arr) => {
        return arr.Response == filterKeys.response;
      });
    }

    if (filterKeys.toUser) {
      dataToBeFiltered = dataToBeFiltered.filter((arr) => {
        return arr.ToUserDetails.UserName.toLowerCase().includes(
          filterKeys.toUser.toLowerCase()
        );
      });
    }

    if (filterKeys.fromUser) {
      dataToBeFiltered = dataToBeFiltered.filter((arr) => {
        return arr.UserDetails.UserName.toLowerCase().includes(
          filterKeys.fromUser.toLowerCase()
        );
      });
    }

    if (filterKeys.fileName) {
      dataToBeFiltered = dataToBeFiltered.filter((arr) => {
        return arr.FileName.toLowerCase().includes(
          filterKeys.fileName.toLowerCase()
        );
      });
    }

    if (filterKeys.product) {
      dataToBeFiltered = dataToBeFiltered.filter((arr) => {
        return arr.Product.toLowerCase().includes(
          filterKeys.product.toLowerCase()
        );
      });
    }

    await setDrData([...dataToBeFiltered]);
    await setDrFilters({ ...filterKeys });
  };
  const queryGenerator = async (filters_: any) => {
    let queryArr = [];
    let queryStr = "";
    let _drAllitems = [];

    // if (filters_.view) {
    // if (filters_.view == "Pending") {
    //   queryArr.push("view-Pending");
    // }
    // if (filters_.view == "Pending edit") {
    //   queryArr.push("view-Pending edit");
    // }
    // if (filters_.view == "Send by me") {
    //   queryArr.push(`FromUser eq ${currentUser["Id"]}`);
    // }
    // if (filters_.view == "Responded by me") {
    //   queryArr.push("view-Responded by me");
    // }
    // if (filters_.view == "Last 30 days") {
    //   queryArr.push("view-Last 30 days");
    // }
    // if (filters_.view == "All") {
    //   queryArr.push("view-All");
    // }
    // }

    // if (filters_.to && filters_.to != "Anyone") {
    //   if (filters_.to == "Me") {
    //     queryArr.push(
    //       `auditTo eq '${props.spcontext.pageContext.user.displayName}'`
    //     );
    //   }
    //   if (filters_.to == "Me or Me Cc'd") {
    //     queryArr.push(
    //       `auditTo eq ${props.spcontext.pageContext.user.displayName} or CcMail eq ${props.spcontext.pageContext.user.Email}`
    //     );
    //   }
    // }

    if (filters_.request && filters_.request != "All") {
      drDropDownOptions.requestOptns.forEach((requestOptn) => {
        if (filters_.request == requestOptn.key) {
          queryArr.push(`auditRequestType eq '${requestOptn.key}'`);
        }
      });
    }

    if (filters_.response && filters_.response != "All") {
      drDropDownOptions.responseOptns.forEach((responseOptn) => {
        if (filters_.response == responseOptn.key) {
          queryArr.push(`auditResponseType eq '${responseOptn.key}'`);
        }
      });
    }

    if (filters_.toUser) {
      queryArr.push(`substringof('${filters_.toUser}',auditTo)`);
    }
    if (filters_.fromUser) {
      queryArr.push(`substringof('${filters_.fromUser}',auditFrom)`);
    }
    if (filters_.fileName) {
      queryArr.push(`substringof('${filters_.fileName}',Title)`);
    }
    if (filters_.product) {
      queryArr.push(`substringof('${filters_.product}',ProductName)`);
    }

    queryStr = queryArr.join(" and ");

    await sharepointWeb.lists
      .getByTitle(drListName)
      .items.select(
        "*",
        "FromUser/Title",
        "FromUser/Id",
        "FromUser/EMail",
        "CcEmail/Title",
        "CcEmail/Id",
        "CcEmail/EMail"
      )
      .expand("FromUser,CcEmail")
      .top(5000)
      .filter(`${queryStr}`)
      .orderBy("Modified", false)
      .get()
      .then(async (items) => {
        items.forEach((item) => {
          let tempCcEmails = [];
          if (item.CcEmailId) {
            item.CcEmail.forEach((email) => {
              peopleList.forEach((people) => {
                if (email.Id == people.ID) {
                  tempCcEmails.push(people);
                }
              });
            });
          }
          _drAllitems.push({
            ID: item.Id ? item.Id : "",
            Request: item.auditRequestType ? item.auditRequestType : "",
            FileName: item.Title ? item.Title : "",
            Sent: item.auditSent
              ? moment(item.auditSent).format("DD/MM/YYYY")
              : "",
            Response: item.auditResponseType ? item.auditResponseType : "",
            UserDetails: {
              UserName: item.FromUser ? item.FromUser.Title : "",
              UserEmail: item.FromUser ? item.FromUser.EMail : "",
              UserId: item.FromUser ? item.FromUser.Id : "",
            },
            User: item.FromUser ? item.FromUser.Title : "",
            RequestComments: item.auditComments ? item.auditComments : "",
            CcEmailIds: item.CcEmailId ? item.CcEmailId : [],
            CcEmails: item.CcEmail ? tempCcEmails : [],
            Modified: item.Modified,
          });
        });
        await setDrData([..._drAllitems]);
      })
      .catch(drErrorFunction);
  };
  const GetUserDetails = (filterText: any) => {
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
  const drReallocateFunction = async () => {
    await sharepointWeb.lists
      .getByTitle(drListName)
      .items.getById(drReviewFormDisplay.selectedItem["ID"])
      .select(
        "*",
        "FromUser/Title",
        "FromUser/Id",
        "FromUser/EMail",
        "CcEmail/Title",
        "CcEmail/Id",
        "CcEmail/EMail"
      )
      .expand("FromUser,CcEmail")
      .get()
      .then(async (item) => {
        const requestCreateData = {
          Title: item.Title ? item.Title : "",
          auditLink: item.auditLink ? item.auditLink : "",
          auditRequestType: item.auditRequestType ? item.auditRequestType : "",
          auditSent: moment().format("MM/DD/yyyy"),
          CcEmailId: item.CcEmailId
            ? { results: item.CcEmailId }
            : { results: [] },
          auditResponseType: "Pending",
          FeedbackRepeated: false,
          auditFrom: item.auditFrom ? item.auditFrom : "",
          FromEmail: item.FromEmail ? item.FromEmail : "",
          FromUserId: item.FromUserId ? item.FromUserId : "",
          auditTo: drReallocateUser["text"] ? drReallocateUser["text"] : "",
          ToEmail: drReallocateUser["secondaryText"]
            ? drReallocateUser["secondaryText"]
            : "",
          ToUserId: drReallocateUser["ID"] ? drReallocateUser["ID"] : "",
          auditComments: `Reallocated from ${item.FromUser.Title} by ${
            currentUser["Name"]
          } originally sent ${moment(new Date()).format(
            "DD/MM/YY HH:mm"
          )} AEST`,
          auditDocLink: item.auditDocLink ? item.auditDocLink : "",
          auditDepartment: item.auditDepartment ? item.auditDepartment : "",
          auditLastResponse: item.auditLastResponse
            ? item.auditLastResponse
            : "",
          auditID: item.auditID ? item.auditID : "",
          auditReplyTo: item.auditReplyTo ? item.auditReplyTo : "",
          auditResponseMeetingRequired: false,
          ResponseAcknowledged: false,
          AnnualPlanID: item.AnnualPlanID ? item.AnnualPlanID : null,
          DeliveryPlanID: item.DeliveryPlanID ? item.DeliveryPlanID : null,
          ProductionBoardID: item.ProductionBoardID
            ? item.ProductionBoardID
            : null,
        };
        const requestUpdateData = {
          auditComments: `Reallocated to ${drReallocateUser["text"]} by ${currentUser["Name"]}`,
          auditResponseType: "Reallocated",
        };

        await drUpdateItem(requestUpdateData, item.Id);
        await drCreateItem(requestCreateData);

        await setDrReallocateUser({});
        await setDrReallocatePopup({
          condition: false,
          allocatedUser: null,
        });
        await setDrReviewFormDisplay({
          condition: false,
          selectedItem: {},
        });
        await setDrReRender(!drReRender);
        await setDrLoader("noLoader");
        await setDrPopup("Reallocate");
        await setTimeout(() => {
          setDrPopup("Close");
        }, 2000);
      })
      .catch(drErrorFunction);
  };
  const drCancelRequestFunction = async () => {
    const requestUpdateData = {
      auditComments: `Reason cancelled: ${drCancelReason} by ${currentUser["Name"]}`,
      auditResponseType: "Cancelled",
    };
    await drUpdateItem(
      requestUpdateData,
      drReviewFormDisplay.selectedItem["ID"]
    );

    await setDrCancelReason("");
    await setDrCancelRequestPopup(false);
    await setDrReviewFormDisplay({
      condition: false,
      selectedItem: {},
    });
    await setDrReRender(!drReRender);
    await setDrLoader("noLoader");
    await setDrPopup("CancelRequest");
    await setTimeout(() => {
      setDrPopup("Close");
    }, 2000);
  };
  const drSubmitFunction = async () => {
    let tempCcEmails = [];
    if (drReviewFormDisplay.selectedItem["CcEmails"].length > 0) {
      drReviewFormDisplay.selectedItem["CcEmails"].forEach((ccEmail) => {
        tempCcEmails.push(ccEmail.ID);
      });
    }
    const requestUpdateData = {
      auditResponseType: drReviewFormOptionDisplay.selectedOption
        ? drReviewFormOptionDisplay.selectedOption
        : "",
      Response_x0020_Comments:
        drReviewFormDisplay.selectedItem["ResponseComments"],
      Rating: drReviewFormOptionDisplay.rating,
      CcEmailId:
        drReviewFormDisplay.selectedItem["CcEmails"].length > 0
          ? { results: tempCcEmails }
          : { results: [] },
      // FeedbackIssues: [drReviewFormOptionDisplay.issuesCategory.issues],
      // FeedbackSeverity:
      //   drReviewFormOptionDisplay.issuesCategory.issuesSeverity,
      // FeedbackRepeated:
      //   drReviewFormOptionDisplay.issuesCategory.issueRepeated,
    };

    await drUpdateItem(
      requestUpdateData,
      drReviewFormDisplay.selectedItem["ID"]
    );

    await setDrReviewFormOptionDisplay({
      condition: false,
      selectedOption: null,
      issuesCategory: {
        issues: "",
        issuesSeverity: "",
        issueRepeated: false,
      },
      rating: 0,
    });
    await setDrReviewFormDisplay({
      condition: false,
      selectedItem: {},
    });
    await setDrReRender(!drReRender);
    await setDrLoader("noLoader");
    await setDrPopup("Submit");
    await setTimeout(() => {
      setDrPopup("Close");
    }, 2000);
  };
  const drSignOffFunction = async () => {
    await sharepointWeb.lists
      .getByTitle(drListName)
      .items.getById(drReviewFormDisplay.selectedItem["ID"])
      .select(
        "*",
        "FromUser/Title",
        "FromUser/Id",
        "FromUser/EMail",
        "CcEmail/Title",
        "CcEmail/Id",
        "CcEmail/EMail"
      )
      .expand("FromUser,CcEmail")
      .get()
      .then(async (item) => {
        let tempCcEmails = [];
        if (drReviewFormDisplay.selectedItem["CcEmails"].length > 0) {
          drReviewFormDisplay.selectedItem["CcEmails"].forEach((ccEmail) => {
            tempCcEmails.push(ccEmail.ID);
          });
        }

        const requestUpdateData = {
          auditResponseType: drReviewFormOptionDisplay.selectedOption
            ? drReviewFormOptionDisplay.selectedOption
            : "",
          Response_x0020_Comments:
            drReviewFormDisplay.selectedItem["ResponseComments"],
          Rating: drReviewFormOptionDisplay.rating,
          CcEmailId:
            drReviewFormDisplay.selectedItem["CcEmails"].length > 0
              ? { results: tempCcEmails }
              : { results: [] },
        };

        await drUpdateItem(
          requestUpdateData,
          drReviewFormDisplay.selectedItem["ID"]
        );
        if (drSignOffOptions.signOffComments) {
          const requestCreateData = {
            Title: item.Title ? item.Title : "",
            auditLink: item.auditLink ? item.auditLink : "",
            auditRequestType: "Sign-off",
            auditSent: moment().format("MM/DD/yyyy"),
            CcEmailId: item.CcEmailId
              ? { results: item.CcEmailId }
              : { results: [] },
            auditResponseType: "Signed Off",
            auditFrom: item.auditFrom ? item.auditFrom : "",
            FromEmail: item.FromEmail ? item.FromEmail : "",
            FromUserId: item.FromUserId ? item.FromUserId : "",
            auditTo: currentUser["Name"] ? currentUser["Name"] : "",
            ToEmail: currentUser["Email"] ? currentUser["Email"] : "",
            ToUserId: currentUser["Id"] ? currentUser["Id"] : "",
            auditComments: `Sign off from Editor as Client Proxy`,
            Response_x0020_Comments: `${drSignOffOptions.signOffComments}`,
            auditDocLink: item.auditDocLink ? item.auditDocLink : "",
            auditDepartment: item.auditDepartment ? item.auditDepartment : "",
            auditLastResponse: item.auditLastResponse
              ? item.auditLastResponse
              : "",
            auditID: item.auditID ? item.auditID : "",
            auditReplyTo: item.auditReplyTo ? item.auditReplyTo : "",
            AnnualPlanID: item.AnnualPlanID ? item.AnnualPlanID : null,
            DeliveryPlanID: item.DeliveryPlanID ? item.DeliveryPlanID : null,
            ProductionBoardID: item.ProductionBoardID
              ? item.ProductionBoardID
              : null,
          };
          await drCreateItem(requestCreateData);
        }
        if (drSignOffOptions.assignTo) {
          const requestCreateData = {
            Title: item.Title ? item.Title : "",
            auditLink: item.auditLink ? item.auditLink : "",
            auditRequestType: "Publish",
            auditSent: moment().format("MM/DD/yyyy"),
            CcEmailId: item.CcEmailId
              ? { results: item.CcEmailId }
              : { results: [] },
            auditResponseType: "Pending",
            auditFrom: item.auditFrom ? item.auditFrom : "",
            FromEmail: item.FromEmail ? item.FromEmail : "",
            FromUserId: item.FromUserId ? item.FromUserId : "",
            auditTo: drSignOffOptions.assignTo["text"]
              ? drSignOffOptions.assignTo["text"]
              : "",
            ToEmail: drSignOffOptions.assignTo["secondaryText"]
              ? drSignOffOptions.assignTo["secondaryText"]
              : "",
            ToUserId: drSignOffOptions.assignTo["ID"]
              ? drSignOffOptions.assignTo["ID"]
              : "",
            auditComments: `${drSignOffOptions.publishRequestComments}`,
            auditDocLink: item.auditDocLink ? item.auditDocLink : "",
            auditDepartment: item.auditDepartment ? item.auditDepartment : "",
            auditLastResponse: item.auditLastResponse
              ? item.auditLastResponse
              : "",
            auditID: item.auditID ? item.auditID : "",
            auditReplyTo: item.auditReplyTo ? item.auditReplyTo : "",
            AnnualPlanID: item.AnnualPlanID ? item.AnnualPlanID : null,
            DeliveryPlanID: item.DeliveryPlanID ? item.DeliveryPlanID : null,
            ProductionBoardID: item.ProductionBoardID
              ? item.ProductionBoardID
              : null,
          };
          await drCreateItem(requestCreateData);
        }
      });

    await setDrReviewFormOptionDisplay({
      condition: false,
      selectedOption: null,
      issuesCategory: {
        issues: "",
        issuesSeverity: "",
        issueRepeated: false,
      },
      rating: 0,
    });
    await setDrSignOffOptions({
      assignTo: null,
      signOffComments: "",
      publishRequestComments: "",
    });
    await setDrSignOffPopup(false);
    await setDrReviewFormDisplay({
      condition: false,
      selectedItem: {},
    });
    await setDrReRender(!drReRender);
    await setDrLoader("noLoader");
    await setDrPopup("SignOff");
    await setTimeout(() => {
      setDrPopup("Close");
    }, 2000);
  };
  const drFixLinkFunction = async () => {
    const requestUpdateData = {
      FixLink: true,
    };
    await drUpdateItem(
      requestUpdateData,
      drReviewFormDisplay.selectedItem["ID"]
    );
    await setDrLoader("noLoader");
    await setDrPopup("FixLink");
    await setTimeout(() => {
      setDrPopup("Close");
    }, 2000);
  };
  const drCreateItem = async (_createData: any) => {
    await sharepointWeb.lists
      .getByTitle("Review Log")
      .items.add(_createData)
      .then(async () => {
        await [];
      })
      .catch(drErrorFunction);
  };
  const drUpdateItem = async (_updateData: any, targetId: number) => {
    await sharepointWeb.lists
      .getByTitle("Review Log")
      .items.getById(targetId)
      .update(_updateData)
      .then(async () => {
        await [];
      })
      .catch(drErrorFunction);
  };
  const drReviewFormOptionHandler = (optionType: string, option: any) => {
    if (optionType == "ResponseComments") {
      let tempSelectedItem = { ...drReviewFormDisplay };
      tempSelectedItem.selectedItem[optionType] = option;
      setDrReviewFormDisplay(tempSelectedItem);
    } else if (optionType == "CcEmails") {
      let tempSelectedItem = { ...drReviewFormDisplay };
      tempSelectedItem.selectedItem[optionType] = [...option];
      setDrReviewFormDisplay(tempSelectedItem);
    } else if (optionType == "rating") {
      let reviewFormOptions = { ...drReviewFormOptionDisplay };
      reviewFormOptions[optionType] = option;
      setDrReviewFormOptionDisplay({ ...reviewFormOptions });
    } else {
      let reviewFormOptions = { ...drReviewFormOptionDisplay };
      reviewFormOptions.issuesCategory[optionType] = option;
      setDrReviewFormOptionDisplay({ ...reviewFormOptions });
    }
  };
  const drSignOffHandler = (key: string, value: any) => {
    let signOffData = { ...drSignOffOptions };
    signOffData[key] = value;
    setDrSignOffOptions(signOffData);
  };
  const SubmitPopup = () => (
    <MessageBar messageBarType={MessageBarType.success} isMultiline={false}>
      Document review has been successfully submitted !!!
    </MessageBar>
  );
  const ReallocatePopup = () => (
    <MessageBar messageBarType={MessageBarType.success} isMultiline={false}>
      Document review has been successfully reallocated !!!
    </MessageBar>
  );
  const CancelRequestPopup = () => (
    <MessageBar messageBarType={MessageBarType.warning} isMultiline={false}>
      Document review - request has been successfully cancelled !!!
    </MessageBar>
  );
  const SignOffPopup = () => (
    <MessageBar messageBarType={MessageBarType.success} isMultiline={false}>
      Document review has been successfully signed off !!!
    </MessageBar>
  );
  const FixLinkPopup = () => (
    <MessageBar messageBarType={MessageBarType.warning} isMultiline={false}>
      We are working on it ,check back later.
    </MessageBar>
  );
  const ErrorPopup = () => (
    <MessageBar messageBarType={MessageBarType.error} isMultiline={false}>
      Something when error, please contact system admin.
    </MessageBar>
  );
  const drErrorFunction = (error: any) => {
    console.log(error);
    setDrPopup("Error");
    setTimeout(() => {
      setDrPopup("Close");
    }, 2000);
  };
  //Function-Section Ends

  useEffect(() => {
    setDrLoader("startUpLoader");
    getAllDRAdmins();
    drGetCurrentUserDetails();
    getAllUsers();
    drGetAllOptions();
  }, [drReRender]);

  return (
    <>
      <div style={{ padding: "5px 10px" }}>
        {drLoader == "startUpLoader" ? <CustomLoader /> : null}
        {/* Popup-Section Starts */}
        <div>
          {drPopup == "Submit"
            ? SubmitPopup()
            : drPopup == "Reallocate"
            ? ReallocatePopup()
            : drPopup == "CancelRequest"
            ? CancelRequestPopup()
            : drPopup == "SignOff"
            ? SignOffPopup()
            : drPopup == "Error"
            ? ErrorPopup()
            : drPopup == "FixLink"
            ? FixLinkPopup()
            : ""}
        </div>
        {/* Popup-Section Ends */}
        {/* Header-Section Starts */}
        <div>
          <div
            className={styles.dpTitle}
            style={{
              justifyContent: "flex-start",
              alignItems: "flex-start",
              marginBottom: "20px",
            }}
          >
            <Label style={{ color: "#000", fontSize: 24, padding: 0 }}>
              Document review
            </Label>
          </div>
          <div
            style={{
              display: "flex",
              alignItems: "center",
              justifyContent: "space-between",
              // marginBottom: 20,
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
              <Toggle
                defaultChecked
                onText="New to old"
                offText="Old to new"
                styles={drToggleStyles}
                onChange={() => {
                  let _allData = [...drData];
                  _allData.reverse();
                  setDrData(_allData);
                }}
              />
              <Label
                style={{
                  marginLeft: 20,
                  marginTop: 5,
                  fontSize: "13px",
                  color: "#323130",
                }}
              >
                Number of records :{" "}
                <b style={{ color: "#038387" }}>{drData.length}</b>
              </Label>
            </div>
            <div
              style={{
                display: "flex",
                alignItems: "center",
                justifyContent: "center",
              }}
            >
              <Label style={{ marginRight: "10px" }}>
                {currentUser["Name"]}
              </Label>
              <Persona
                size={PersonaSize.size24}
                presence={PersonaPresence.none}
                imageUrl={
                  "/_layouts/15/userphoto.aspx?size=S&username=" +
                  `${currentUser["Email"]}`
                }
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
              <Label styles={drLabelStyles}>View</Label>
              <Dropdown
                placeholder="Select an option"
                styles={drDropdownStyles}
                options={drDropDownOptions.viewOptns}
                dropdownWidth={"auto"}
                onChange={(e, option: any) => {
                  drhandleFilters("view", option["key"]);
                }}
                selectedKey={drFilters.view}
              />
            </div>
            <div>
              <Label styles={drLabelStyles}>To</Label>
              <Dropdown
                placeholder="Select an option"
                styles={drDropdownStyles}
                options={drDropDownOptions.toOptns}
                dropdownWidth={"auto"}
                onChange={(e, option: any) => {
                  drhandleFilters("to", option["key"]);
                }}
                selectedKey={drFilters.to}
              />
            </div>
            <div>
              <Label styles={drLabelStyles}>Request</Label>
              <Dropdown
                placeholder="Select an option"
                styles={drDropdownStyles}
                options={drDropDownOptions.requestOptns}
                dropdownWidth={"auto"}
                onChange={(e, option: any) => {
                  drhandleFilters("request", option["key"]);
                }}
                selectedKey={drFilters.request}
              />
            </div>
            <div>
              <Label styles={drLabelStyles}>Response</Label>
              <Dropdown
                placeholder="Select an option"
                styles={drDropdownStyles}
                options={drDropDownOptions.responseOptns}
                dropdownWidth={"auto"}
                onChange={(e, option: any) => {
                  drhandleFilters("response", option["key"]);
                }}
                selectedKey={drFilters.response}
              />
            </div>
            <div>
              <Label styles={drLabelStyles}>Filter (To user)</Label>
              <SearchBox
                styles={drSearchBoxStyles}
                value={drFilters.toUser}
                onChange={(e, value) => {
                  drhandleFilters("toUser", value);
                }}
              />
            </div>
            <div>
              <Label styles={drLabelStyles}>Filter (From user)</Label>
              {/* <NormalPeoplePicker
                styles={drModalBoxPP}
                onResolveSuggestions={GetUserDetails}
                itemLimit={1}
                onChange={(selectedUser) => {
                  selectedUser.length > 0
                    ? console.log(selectedUser[0].text)
                    : null;
                }}
              /> */}
              <SearchBox
                styles={drSearchBoxStyles}
                value={drFilters.fromUser}
                onChange={(e, value) => {
                  drhandleFilters("fromUser", value);
                }}
              />
            </div>
            <div>
              <Label styles={drLabelStyles}>Filter (Filename)</Label>
              <SearchBox
                styles={drSearchBoxStyles}
                value={drFilters.fileName}
                onChange={(e, value) => {
                  drhandleFilters("fileName", value);
                }}
              />
            </div>
            <div>
              <Label styles={drLabelStyles}>Filter (product)</Label>
              <SearchBox
                styles={drSearchBoxStyles}
                value={drFilters.product}
                onChange={(e, value) => {
                  drhandleFilters("product", value);
                }}
              />
            </div>
            <div>
              <Icon
                iconName="Refresh"
                className={drIconStyleClass.refresh}
                onClick={() => {
                  let tempResetFilters = {
                    view: "Pending",
                    to: "Me",
                    request: "All",
                    response: "All",
                    toUser: "",
                    fromUser: "",
                    fileName: "",
                    product: "",
                  };
                  filterItems(tempResetFilters);
                }}
              />
            </div>
          </div>
        </div>
        {/* Filter-Section Ends */}
        {/* Body-Section Starts */}
        <div style={{ display: "flex" }}>
          {/* DetailList-Section Starts */}
          {drData.length > 0 ? (
            <div>
              <DetailsList
                items={drData}
                columns={drColumns}
                styles={drDetailsListStyles}
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
          {/* Form-Section Starts */}
          <div>
            {drReviewFormDisplay.condition ? (
              <div
                style={{
                  // width: 725,
                  width: 800,
                  height: 460,
                  marginTop: 16,
                  overflowX: "hidden",
                  overflowY: "scroll",
                }}
                className={styles.requestReviewPanel}
              >
                <div
                  style={{
                    height: 500,
                    // width: 675,
                    width: 750,
                    marginLeft: 20,
                  }}
                >
                  <div
                    style={{
                      display: "flex",
                      justifyContent: "center",
                      alignItems: "center",
                      position: "relative",
                    }}
                  >
                    <span className={generalStyles.titleLabel}>
                      {`Request 
                      ${drReviewFormDisplay.selectedItem[
                        "Request"
                      ].toLowerCase()}`}
                    </span>
                    <span
                      style={{
                        color: "#959595",
                        position: "absolute",
                        right: "10px",
                        top: "0",
                        fontWeight: "500",
                      }}
                    >
                      {moment(drReviewFormDisplay.selectedItem["Sent"]).format(
                        "DD/MM/YYYY h:mm A"
                      )}
                    </span>
                  </div>
                  <div className={styles.drRequestFormBtnSection}>
                    <div style={{ display: "flex" }}>
                      <a
                        href={`${drReviewFormDisplay.selectedItem["Link"]}?web=1`}
                        data-interception="off"
                        target="_blank"
                      >
                        <button className={styles.openFileBtn}>
                          Open file
                        </button>
                      </a>
                      <button
                        className={styles.fixLinkBtn}
                        onClick={() => {
                          drFixLinkFunction();
                          setDrLoader("FixLink");
                        }}
                      >
                        {drLoader == "FixLink" ? <Spinner /> : "Fix Link"}
                      </button>
                    </div>
                    <div style={{ display: "flex" }}>
                      <button
                        className={
                          drReviewFormDisplay.selectedItem["Response"] ==
                          "Pending"
                            ? styles.reallocateBtn
                            : styles.disableBtn
                        }
                        onClick={() => {
                          drReviewFormDisplay.selectedItem["Response"] ==
                          "Pending"
                            ? setDrReallocatePopup({
                                condition: true,
                                allocatedUser:
                                  drReviewFormDisplay.selectedItem[
                                    "UserDetails"
                                  ].UserId,
                              })
                            : "";
                        }}
                      >
                        Reallocate
                      </button>
                      <button
                        className={
                          drReviewFormDisplay.selectedItem["Response"] ==
                          "Pending"
                            ? styles.cancelRequestBtn
                            : styles.disableBtn
                        }
                        onClick={() => {
                          drReviewFormDisplay.selectedItem["Response"] ==
                          "Pending"
                            ? setDrCancelRequestPopup(true)
                            : "";
                        }}
                      >
                        Cancel Request
                      </button>
                    </div>
                  </div>
                  <div
                    style={{ display: "flex", justifyContent: "space-between" }}
                  >
                    <div className={generalStyles.inputField}>
                      <label className={generalStyles.inputLabel}>File</label>
                      <span className={generalStyles.inputValue}>
                        {drReviewFormDisplay.selectedItem["FileName"]}
                      </span>
                    </div>
                    <div className={generalStyles.inputField}>
                      <label className={generalStyles.inputLabel}>
                        Current Response
                      </label>
                      <span className={generalStyles.inputValue}>
                        {drReviewFormDisplay.selectedItem["Response"]}
                      </span>
                    </div>
                    <div className={generalStyles.inputField}>
                      <label className={generalStyles.inputLabel}>
                        From User
                      </label>
                      <span className={generalStyles.inputValue}>
                        {
                          drReviewFormDisplay.selectedItem["UserDetails"]
                            .UserName
                        }
                      </span>
                    </div>
                  </div>
                  <div
                    className={generalStyles.inputField}
                    style={{ display: "flex", justifyContent: "space-between" }}
                  >
                    <div>
                      <label className={generalStyles.inputLabel}>
                        Request Comments
                      </label>
                      <div className={styles.reviewDesc} style={{}}>
                        {drReviewFormDisplay.selectedItem["RequestComments"]}
                      </div>
                    </div>
                    {drReviewFormDisplay.selectedItem["Response"] !=
                    "Pending" ? (
                      <div>
                        <label className={generalStyles.inputLabel}>
                          Rating
                        </label>
                        <Rating
                          max={4}
                          rating={drReviewFormDisplay.selectedItem["Rating"]}
                          disabled={true}
                          style={{ width: 120 }}
                          size={RatingSize.Large}
                        />
                      </div>
                    ) : (
                      ""
                    )}
                    {drReviewFormDisplay.selectedItem["Response"] !=
                    "Pending" ? (
                      <div>
                        <label className={generalStyles.inputLabel}>
                          Repeated Issues
                        </label>
                        <Toggle
                          onText="Yes"
                          offText="No"
                          checked={
                            drReviewFormDisplay.selectedItem["RepeatedIssue"]
                          }
                          disabled={true}
                          styles={{ root: { marginTop: "15px" } }}
                        />
                      </div>
                    ) : (
                      ""
                    )}
                  </div>
                  {drReviewFormDisplay.selectedItem["Response"] == "Pending" ? (
                    <div
                      className={generalStyles.inputField}
                      style={{
                        display: "flex",
                        alignItems: "center",
                        justifyContent: "flex-start",
                      }}
                    >
                      <div>
                        <label className={generalStyles.inputLabel}>
                          Response
                        </label>
                        <Dropdown
                          placeholder="Select an option"
                          selectedKey={drReviewFormOptionDisplay.selectedOption}
                          options={
                            drReviewFormDisplay.selectedItem["Request"] ==
                            "Report"
                              ? [
                                  { key: "Select", text: "Select" },
                                  { key: "Completed", text: "Completed" },
                                ]
                              : drReviewFormDisplay.selectedItem["Request"] ==
                                "Review"
                              ? [
                                  { key: "Select", text: "Select" },
                                  { key: "Feedback", text: "Feedback" },
                                  { key: "Returned", text: "Returned" },
                                  { key: "Endorsed", text: "Endorsed" },
                                  { key: "Signed Off", text: "Signed Off" },
                                ]
                              : drReviewFormDisplay.selectedItem["Request"] ==
                                "Initial Edit"
                              ? [
                                  { key: "Select", text: "Select" },
                                  { key: "Edited", text: "Edited" },
                                  { key: "Returned", text: "Returned" },
                                ]
                              : drReviewFormDisplay.selectedItem["Request"] ==
                                "Assemble"
                              ? [
                                  { key: "Select", text: "Select" },
                                  { key: "Assembled", text: "Assembled" },
                                  { key: "Returned", text: "Returned" },
                                ]
                              : drReviewFormDisplay.selectedItem["Request"] ==
                                "Add Images"
                              ? [
                                  { key: "Select", text: "Select" },
                                  { key: "Inserted", text: "Inserted" },
                                  { key: "Returned", text: "Returned" },
                                ]
                              : drReviewFormDisplay.selectedItem["Request"] ==
                                "Publish"
                              ? [
                                  { key: "Select", text: "Select" },
                                  { key: "Published", text: "Published" },
                                  { key: "Returned", text: "Returned" },
                                ]
                              : drReviewFormDisplay.selectedItem["Request"] ==
                                "Final Edit"
                              ? [
                                  { key: "Select", text: "Select" },
                                  { key: "Edited", text: "Edited" },
                                  { key: "Returned", text: "Returned" },
                                ]
                              : drReviewFormDisplay.selectedItem["Request"] ==
                                "Sign-off"
                              ? [
                                  { key: "Select", text: "Select" },
                                  { key: "Signed Off", text: "Signed Off" },
                                  { key: "Returned", text: "Returned" },
                                ]
                              : [{ key: "Select", text: "Select" }]
                          }
                          dropdownWidth={"auto"}
                          styles={drReviewFormDropDownStyles}
                          onChange={(e, option) => {
                            option.key != "Select"
                              ? setDrReviewFormOptionDisplay({
                                  condition: true,
                                  selectedOption: option.key,
                                  issuesCategory: {
                                    issues: "",
                                    issuesSeverity: "",
                                    issueRepeated: false,
                                  },
                                  rating: 4,
                                })
                              : setDrReviewFormOptionDisplay({
                                  condition: false,
                                  selectedOption: null,
                                  issuesCategory: {
                                    issues: "",
                                    issuesSeverity: "",
                                    issueRepeated: false,
                                  },
                                  rating: 0,
                                });
                          }}
                        />
                      </div>

                      <div
                        // className={generalStyles.inputField}
                        style={{
                          display: "flex",
                          alignItems: "center",
                          justifyContent: "flex-start",
                        }}
                      >
                        {drReviewFormOptionDisplay.selectedOption ? (
                          <div style={{ marginLeft: "30px" }}>
                            <label className={generalStyles.inputLabel}>
                              Rating
                            </label>
                            <div
                              style={{
                                display: "flex",
                                justifyContent: "flex-start",
                              }}
                            >
                              <Rating
                                max={4}
                                rating={drReviewFormOptionDisplay.rating}
                                styles={
                                  drReviewFormOptionDisplay.rating == 4
                                    ? { ratingStarFront: { color: "#00D100" } }
                                    : drReviewFormOptionDisplay.rating == 3
                                    ? { ratingStarFront: { color: "#FFFF00" } }
                                    : drReviewFormOptionDisplay.rating == 2
                                    ? { ratingStarFront: { color: "#D18700" } }
                                    : { ratingStarFront: { color: "#D10000" } }
                                }
                                style={{ width: 120 }}
                                size={RatingSize.Large}
                                onChange={(e, value) => {
                                  drReviewFormOptionHandler("rating", value);
                                }}
                              />
                              <Label
                                style={{
                                  width: 200,
                                  marginRight: 20,
                                  marginTop: 2,
                                  fontSize: 15,
                                }}
                                styles={
                                  drReviewFormOptionDisplay.rating == 4
                                    ? { root: { color: "#00D100" } }
                                    : drReviewFormOptionDisplay.rating == 3
                                    ? { root: { color: "#FFFF00" } }
                                    : drReviewFormOptionDisplay.rating == 2
                                    ? { root: { color: "#D18700" } }
                                    : { root: { color: "#D10000" } }
                                }
                              >
                                {drReviewFormOptionDisplay.rating == 4
                                  ? " - Exceed"
                                  : drReviewFormOptionDisplay.rating == 3
                                  ? " - Good"
                                  : drReviewFormOptionDisplay.rating == 2
                                  ? " - Average"
                                  : " - Needs improvement"}
                              </Label>
                            </div>
                          </div>
                        ) : (
                          // : drReviewFormOptionDisplay.selectedOption ? (
                          //   <div>
                          //     <label className={generalStyles.inputLabel}>Issues Severity</label>
                          //     <Dropdown
                          //       placeholder="Find Items"
                          //       options={[
                          //         { key: "Major", text: "Major" },
                          //         { key: "Moderate", text: "Moderate" },
                          //         { key: "Minor", text: "Minor" },
                          //         { key: "None", text: "None" },
                          //       ]}
                          //       dropdownWidth={"auto"}
                          //       styles={drReviewFormDropDownStyles}
                          //       selectedKey={"None"}
                          //       onChange={(e, option) => {
                          //         drReviewFormOptionHandler(
                          //           "issuesSeverity",
                          //           option["key"]
                          //         );
                          //       }}
                          //     />
                          //   </div>
                          // )
                          ""
                        )}

                        {drReviewFormOptionDisplay.condition ? (
                          <div className={generalStyles.inputField}>
                            <label className={generalStyles.inputLabel}>
                              Repeated Issues
                            </label>
                            <Toggle
                              onText="Yes"
                              offText="No"
                              // style={{ marginTop: 15 }}
                              styles={{ root: { marginTop: "15px" } }}
                              onChange={() => {
                                drReviewFormOptionHandler(
                                  "issueRepeated",
                                  !drReviewFormOptionDisplay.issuesCategory
                                    .issueRepeated
                                );
                              }}
                            />
                          </div>
                        ) : (
                          ""
                        )}
                      </div>

                      {/* {drReviewFormOptionDisplay.selectedOption &&
                      drReviewFormOptionDisplay.selectedOption != "Completed" &&
                      drReviewFormOptionDisplay.selectedOption != "Returned" &&
                      drReviewFormOptionDisplay.selectedOption !=
                        "Signed Off" ? (
                        <div>
                          <Label>Issues</Label>
                          <Dropdown
                            placeholder="Find Items"
                            options={[
                              {
                                key: "Style/Formatting",
                                text: "Style/Formatting",
                              },
                              {
                                key: "Content Quality",
                                text: "Content Quality",
                              },
                              { key: "Incomplete", text: "Incomplete" },
                            ]}
                            dropdownWidth={"auto"}
                            styles={drReviewFormDropDownStyles}
                            onChange={(e, option) => {
                              drReviewFormOptionHandler(
                                "issues",
                                option["key"]
                              );
                            }}
                          />
                        </div>
                      ) : (
                        ""
                      )} */}
                    </div>
                  ) : (
                    ""
                  )}
                  {/* <div
                    className={generalStyles.inputField}
                    style={{
                      display: "flex",
                      alignItems: "center",
                      justifyContent: "flex-start",
                    }}
                  >
                    {drReviewFormOptionDisplay.selectedOption ? (
                      // &&
                      //   drReviewFormOptionDisplay.selectedOption == "Completed" ||
                      // drReviewFormOptionDisplay.selectedOption == "Returned" ||
                      // drReviewFormOptionDisplay.selectedOption == "Signed Off"
                      <div>
                        <label className={generalStyles.inputLabel}>
                          Rating
                        </label>
                        <Rating
                          max={4}
                          style={{ width: 345 }}
                          defaultRating={drReviewFormOptionDisplay.rating}
                          size={RatingSize.Large}
                          ariaLabel="Large stars"
                          onChange={(e, value) => {
                            drReviewFormOptionHandler("rating", value);
                          }}
                        />
                      </div>
                    ) : (
                      // : drReviewFormOptionDisplay.selectedOption ? (
                      //   <div>
                      //     <label className={generalStyles.inputLabel}>Issues Severity</label>
                      //     <Dropdown
                      //       placeholder="Find Items"
                      //       options={[
                      //         { key: "Major", text: "Major" },
                      //         { key: "Moderate", text: "Moderate" },
                      //         { key: "Minor", text: "Minor" },
                      //         { key: "None", text: "None" },
                      //       ]}
                      //       dropdownWidth={"auto"}
                      //       styles={drReviewFormDropDownStyles}
                      //       selectedKey={"None"}
                      //       onChange={(e, option) => {
                      //         drReviewFormOptionHandler(
                      //           "issuesSeverity",
                      //           option["key"]
                      //         );
                      //       }}
                      //     />
                      //   </div>
                      // )
                      ""
                    )}
                    {drReviewFormOptionDisplay.condition ? (
                      <div className={generalStyles.inputField}>
                        <label className={generalStyles.inputLabel}>
                          Repeated Issues
                        </label>
                        <Toggle
                          onText="On"
                          offText="Off"
                          styles={drToggleStyles}
                          onChange={() => {
                            drReviewFormOptionHandler(
                              "issueRepeated",
                              !drReviewFormOptionDisplay.issuesCategory
                                .issueRepeated
                            );
                          }}
                        />
                      </div>
                    ) : (
                      ""
                    )}
                  </div> */}
                  <div className={generalStyles.inputField}>
                    <label className={generalStyles.inputLabel}>Cc Email</label>
                    <NormalPeoplePicker
                      disabled={
                        drReviewFormDisplay.selectedItem["Response"] ==
                        "Pending"
                          ? false
                          : true
                      }
                      inputProps={{
                        placeholder:
                          drReviewFormDisplay.selectedItem["Response"] ==
                          "Pending"
                            ? "Find People"
                            : "",
                      }}
                      styles={drReviewFormPP}
                      onResolveSuggestions={GetUserDetails}
                      selectedItems={
                        drReviewFormDisplay.selectedItem["CcEmails"]
                      }
                      onChange={(selectedUser) => {
                        drReviewFormOptionHandler("CcEmails", selectedUser);
                      }}
                    />
                  </div>
                  <div
                    className={generalStyles.inputField}
                    style={{ marginTop: "22px", marginBottom: "60px" }}
                  >
                    <label
                      className={generalStyles.inputLabel}
                      style={{ margin: "10px 10px 10px 0" }}
                    >
                      Response Comments
                    </label>
                    {/* <TextField
                      multiline
                      resizable={false}
                      value={
                        drReviewFormDisplay.selectedItem["ResponseComments"]
                      }
                      disabled={
                        drReviewFormDisplay.selectedItem["Response"] ==
                        "Pending"
                          ? false
                          : true
                      }
                      onChange={(e, value: string) => {
                        drReviewFormOptionHandler("ResponseComments", value);
                      }}
                      style={{ height: "100px", width: "100%" }}
                    /> */}
                    <ReactQuill
                      theme="snow"
                      modules={modules}
                      formats={formats}
                      readOnly={
                        drReviewFormDisplay.selectedItem["Response"] ==
                        "Pending"
                          ? false
                          : true
                      }
                      value={
                        drReviewFormDisplay.selectedItem["ResponseComments"]
                          ? drReviewFormDisplay.selectedItem["ResponseComments"]
                          : ""
                      }
                      onChange={(e) => {
                        console.log(e);
                        drReviewFormOptionHandler("ResponseComments", e);
                      }}
                      style={{
                        height: 100,
                        width: 750,
                      }}
                    ></ReactQuill>
                  </div>
                  <div>
                    <div
                      className={`${styles.drReviewSubmitBtnSection} ${generalStyles.inputField}`}
                    >
                      {(drReviewFormOptionDisplay.selectedOption == "Edited" ||
                        drReviewFormOptionDisplay.selectedOption ==
                          "Signed Off") &&
                      (documentReviewAdmins.some(
                        (admin) =>
                          admin.text.toLowerCase() ==
                          currentUser["Name"].toLowerCase()
                      ) == true ||
                        currentUser["Email"] ==
                          "nprince@goodtogreatschools.org.au") ? (
                        <button
                          className={styles.drRequestFormPublishBtn}
                          onClick={() => {
                            setDrSignOffPopup(true);
                          }}
                        >
                          Sign Off and Publish
                        </button>
                      ) : (
                        ""
                      )}
                      <button
                        className={
                          drReviewFormOptionDisplay.selectedOption
                            ? styles.drRequestFormSubmitBtn
                            : styles.drRequestFormBtnDisabled
                        }
                        onClick={() => {
                          if (drReviewFormOptionDisplay.selectedOption) {
                            drSubmitFunction();
                            setDrLoader("Submit");
                          }
                        }}
                      >
                        {drLoader == "Submit" ? <Spinner /> : "Submit"}
                      </button>
                    </div>
                  </div>
                </div>
              </div>
            ) : (
              <>
                {drData.length > 0 ? (
                  <div style={{ marginLeft: 360, marginTop: 250 }}>
                    <label
                      // className={generalStyles.inputLabel}
                      style={{
                        color: "#959595 ",
                        display: "block",
                        fontWeight: "500",
                        margin: "5px 0",
                      }}
                    >
                      No Item Selected !!!
                    </label>
                  </div>
                ) : (
                  ""
                )}
              </>
            )}
          </div>
          {/* Form-Section Ends */}
          {/* Popup-Section Starts */}
          {drReallocatePopup.condition ? (
            <Modal
              isOpen={drReallocatePopup.condition}
              isBlocking={true}
              styles={drModalStyles}
            >
              <div>
                <Label className={styles.drPopupLabel}>Reallocate</Label>
                <div className={styles.drPopupDescription}>
                  This will close this request and allocate a new request for
                  the selected user
                </div>
                <div>
                  <NormalPeoplePicker
                    styles={drModalBoxPP}
                    onResolveSuggestions={GetUserDetails}
                    itemLimit={1}
                    defaultSelectedItems={peopleList.filter((people) => {
                      return people.ID == drReallocatePopup.allocatedUser;
                    })}
                    onChange={(selectedUser) => {
                      setDrReallocateUser(selectedUser[0]);
                    }}
                  />
                </div>
                <div className={styles.drPopupButtonSection}>
                  <button
                    className={
                      drReallocateUser
                        ? styles.successBtnActive
                        : styles.successBtnInActive
                    }
                    onClick={() => {
                      if (drReallocateUser) {
                        setDrLoader("Reallocate");
                        drReallocateFunction();
                      }
                    }}
                  >
                    {drLoader == "Reallocate" ? <Spinner /> : "Reallocate"}
                  </button>
                  <button
                    className={styles.closeBtn}
                    onClick={() => {
                      setDrReallocatePopup({
                        condition: false,
                        allocatedUser: null,
                      });
                      setDrReallocateUser({});
                    }}
                  >
                    Close
                  </button>
                </div>
              </div>
            </Modal>
          ) : (
            ""
          )}
          {drCancelRequestPopup ? (
            <Modal
              isOpen={drCancelRequestPopup}
              isBlocking={true}
              styles={drModalStyles}
            >
              <div>
                <Label className={styles.drPopupLabel}>Confirmation</Label>
                <div className={styles.drPopupDescription}>
                  This will cancel the request and remove from the persons
                  review log. Enter your reason to cancel.
                </div>
                <TextField
                  styles={drModalTextFields}
                  onChange={(e, value) => {
                    setDrCancelReason(value);
                  }}
                  placeholder="Reason for cancelling"
                />
                <div className={styles.drPopupDescription}>
                  Do you wish to proceed?
                </div>
                <div className={styles.drPopupButtonSection}>
                  <button
                    className={
                      drCancelReason
                        ? styles.successBtnActive
                        : styles.successBtnInActive
                    }
                    onClick={() => {
                      if (drCancelReason) {
                        setDrLoader("cancelRequest");
                        drCancelRequestFunction();
                      }
                    }}
                  >
                    {drLoader == "cancelRequest" ? <Spinner /> : "Yes"}
                  </button>
                  <button
                    className={styles.closeBtn}
                    onClick={() => {
                      setDrCancelRequestPopup(false);
                      setDrCancelReason("");
                    }}
                  >
                    Close
                  </button>
                </div>
              </div>
            </Modal>
          ) : (
            ""
          )}
          {drSignOffPopup ? (
            <Modal
              isOpen={drSignOffPopup}
              isBlocking={true}
              styles={drModalStyles}
            >
              <div>
                <Label className={styles.drPopupLabel}>
                  Sign Off and pubblish
                </Label>
                <div className={styles.drPopupDescription}>
                  This will save your current response and then sign off and
                  publish (if selected)
                </div>
                <NormalPeoplePicker
                  inputProps={{
                    placeholder:
                      "Assign Publish to, leave blank to not publish",
                  }}
                  styles={drModalBoxPP}
                  onResolveSuggestions={GetUserDetails}
                  itemLimit={1}
                  onChange={(selectedUser) => {
                    drSignOffHandler("assignTo", selectedUser[0]);
                  }}
                />
                <TextField
                  styles={drModalTextFields}
                  defaultValue={drSignOffOptions.signOffComments}
                  onChange={(e, value) => {
                    drSignOffHandler("signOffComments", value);
                  }}
                  placeholder="Sign Off Comments"
                />
                <TextField
                  styles={drModalTextFields}
                  defaultValue={drSignOffOptions.publishRequestComments}
                  onChange={(e, value) => {
                    drSignOffHandler("publishRequestComments", value);
                  }}
                  placeholder="Publish request comments (if publishing"
                />
                <div className={styles.drPopupButtonSection}>
                  <button
                    className={
                      drSignOffOptions.assignTo ||
                      drSignOffOptions.signOffComments ||
                      drSignOffOptions.publishRequestComments
                        ? styles.successBtnActive
                        : styles.successBtnInActive
                    }
                    onClick={() => {
                      if (
                        drSignOffOptions.assignTo ||
                        drSignOffOptions.signOffComments ||
                        drSignOffOptions.publishRequestComments
                      ) {
                        setDrLoader("signOff");
                        drSignOffFunction();
                      }
                    }}
                  >
                    {drLoader == "signOff" ? <Spinner /> : "Yes"}
                  </button>
                  <button
                    className={styles.closeBtn}
                    onClick={() => {
                      setDrSignOffPopup(false);
                      setDrSignOffOptions({
                        assignTo: null,
                        signOffComments: "",
                        publishRequestComments: "",
                      });
                    }}
                  >
                    Close
                  </button>
                </div>
              </div>
            </Modal>
          ) : (
            ""
          )}
          {/* Popup-Section Ends */}
        </div>
        {/* Body-Section Ends */}
      </div>
    </>
  );
};

export default DocumentReview;
