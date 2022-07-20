import * as React from "react";
import { useState, useEffect } from "react";
import * as moment from "moment";
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
import "../ExternalRef/styleSheets/activityStyles.css";
import styles from "./InnovationHubIntranet.module.scss";
import { IDetailsListStyles } from "office-ui-fabric-react";
import { mergeStyles, mergeStyleSets } from "@fluentui/react/lib/Styling";

const ActivityTemplate = (props: any) => {
  const sharepointWeb = Web(props.URL);

  const ATDrpDwnOptns = {
    Project: [{ key: "All", text: "All" }],
    Type: [{ key: "All", text: "All" }],
    Area: [{ key: "All", text: "All" }],
    Product: [{ key: "All", text: "All" }],
  };
  const ATFilterKeys = {
    Project: "All",
    Type: "All",
    Area: "All",
    Product: "All",
  };
  const _dpColumns = [
    {
      key: "Column1",
      name: "Template",
      fieldName: "Title",
      minWidth: 200,
      maxWidth: 200,
    },
    {
      key: "Column2",
      name: "Project",
      fieldName: "Project",
      minWidth: 200,
      maxWidth: 200,
    },
    {
      key: "Column3",
      name: "Type",
      fieldName: "Type",
      minWidth: 200,
      maxWidth: 200,
    },
    {
      key: "Column4",
      name: "Area",
      fieldName: "Area",
      minWidth: 200,
      maxWidth: 200,
    },
    {
      key: "Column5",
      name: "Product",
      fieldName: "Product",
      minWidth: 200,
      maxWidth: 200,
    },
    {
      key: "Column6",
      name: "Code",
      fieldName: "Code",
      minWidth: 120,
      maxWidth: 120,
    },
    {
      key: "Column7",
      name: "Action",
      fieldName: "",
      minWidth: 50,
      maxWidth: 50,
      onRender: (item) => (
        <>
          <Icon
            iconName="Edit"
            className={ATIconStyleClass.edit}
            onClick={() => {}}
          />
          <Icon
            iconName="Delete"
            className={ATIconStyleClass.delete}
            onClick={() => {}}
          />
        </>
      ),
    },
  ];

  const ATIconStyle = mergeStyles({
    fontSize: 17,
    height: 14,
    width: 17,
    cursor: "pointer",
  });
  const ATIconStyleClass = mergeStyleSets({
    link: [{ color: "#2392B2", margin: "0" }, ATIconStyle],
    delete: [{ color: "#CB1E06", margin: "0 7px " }, ATIconStyle],
    edit: [{ color: "#2392B2", margin: "0 7px 0 0" }, ATIconStyle],
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

  // detailslist
  const gridStyles: Partial<IDetailsListStyles> = {
    root: {
      selectors: {
        "& [role=grid]": {
          display: "flex",
          flexDirection: "column",
          ".ms-DetailsRow-fields": {
            alignItems: "center",
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
  const ATBigiconStyleClass = mergeStyleSets({
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
  const ATbuttonStyle = mergeStyles({
    textAlign: "center",
    borderRadius: "2px",
  });
  const ATbuttonStyleClass = mergeStyleSets({
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
      ATbuttonStyle,
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
      ATbuttonStyle,
    ],
  });
  const ATlabelStyles = mergeStyleSets({
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
  const ATdropdownStyles: Partial<IDropdownStyles> = {
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

  // Functions
  const getProductList = () => {
    let _ATdata = [];
    sharepointWeb.lists
      .getByTitle("Activity Template")
      .items.get()
      .then(async (items) => {
        items.forEach((item) => {
          _ATdata.push({
            ID: item.Id,
            Title: item.Title,
            Project: item.Project,
            Type: item.Types,
            Area: item.Area,
            Product: item.Product,
            Code: item.Code,
          });
        });
        setATData([..._ATdata]);
        setATMasterData([..._ATdata]);
        reloadFilterOptions([..._ATdata]);
      })
      .catch(dpErrorFunction);
  };
  const reloadFilterOptions = (data) => {
    let tempArrReload = data;

    tempArrReload.forEach((at) => {
      if (
        ATDrpDwnOptns.Project.findIndex((prj) => {
          return prj.key == at.Project;
        }) == -1
      ) {
        ATDrpDwnOptns.Project.push({
          key: at.Project,
          text: at.Project,
        });
      }
      if (
        ATDrpDwnOptns.Type.findIndex((type) => {
          return type.key == at.Type;
        }) == -1
      ) {
        ATDrpDwnOptns.Type.push({
          key: at.Type,
          text: at.Type,
        });
      }
      if (
        ATDrpDwnOptns.Area.findIndex((area) => {
          return area.key == at.Area;
        }) == -1
      ) {
        ATDrpDwnOptns.Area.push({
          key: at.Area,
          text: at.Area,
        });
      }
      if (
        ATDrpDwnOptns.Product.findIndex((prd) => {
          return prd.key == at.Product;
        }) == -1
      ) {
        ATDrpDwnOptns.Product.push({
          key: at.Product,
          text: at.Product,
        });
      }
    });
    setATDropDownOptions(ATDrpDwnOptns);
  };
  const dpErrorFunction = (error) => {
    // setdpLoader(false);
    console.log(error);
    // if (dpData.length > 0) {
    //   setdpPopup("Error");
    //   setTimeout(() => {
    //     setdpPopup("Close");
    //   }, 2000);
    // }
  };

  // Onchange and Filters
  const ATListFilter = (key, option) => {
    let tempArr = [...ATMasterData];
    let tempDpFilterKeys = { ...ATFilterOptions };
    tempDpFilterKeys[`${key}`] = option;

    if (tempDpFilterKeys.Project != "All") {
      tempArr = tempArr.filter((arr) => {
        return arr.Project == tempDpFilterKeys.Project;
      });
    }
    if (tempDpFilterKeys.Type != "All") {
      tempArr = tempArr.filter((arr) => {
        return arr.Type == tempDpFilterKeys.Type;
      });
    }
    if (tempDpFilterKeys.Area != "All") {
      tempArr = tempArr.filter((arr) => {
        return arr.Area == tempDpFilterKeys.Area;
      });
    }
    if (tempDpFilterKeys.Product != "All") {
      tempArr = tempArr.filter((arr) => {
        return arr.Product == tempDpFilterKeys.Product;
      });
    }

    setATData([...tempArr]);
    setATFilterOptions({ ...tempDpFilterKeys });
  };

  //Use State
  const [ATReRender, setATReRender] = useState(false);
  const [ATData, setATData] = useState([]);
  const [ATMasterData, setATMasterData] = useState([]);
  const [ATDropDownOptions, setATDropDownOptions] = useState(ATDrpDwnOptns);
  const [ATFilterOptions, setATFilterOptions] = useState(ATFilterKeys);
  // Use Effect
  useEffect(() => {
    getProductList();
  }, [ATReRender]);

  return (
    <div style={{ padding: "5px 15px" }}>
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
            className={ATBigiconStyleClass.ChevronLeftMed}
            onClick={() => {}}
          />
          <Label style={{ color: "#000", fontSize: 24, padding: 0 }}>
            Activity planner
          </Label>
        </div>
      </div>
      <div
        style={{
          display: "flex",
          alignItems: "center",
          justifyContent: "left",
          marginBottom: "5px",
        }}
      >
        <PrimaryButton
          text="Add"
          className={ATbuttonStyleClass.buttonPrimary}
          onClick={(_) => {}}
        />
      </div>
      <div
        style={{
          display: "flex",
          alignItems: "center",
          justifyContent: "space-between",
          paddingBottom: "10px",
        }}
      >
        <div className={styles.ddSection}>
          <div>
            <Label className={ATlabelStyles.inputLabels}>Project</Label>
            <Dropdown
              selectedKey={ATFilterOptions.Project}
              placeholder="Select an option"
              options={ATDropDownOptions.Project}
              styles={ATdropdownStyles}
              onChange={(e, option: any) => {
                ATListFilter("Project", option["key"]);
              }}
            />
          </div>
          <div>
            <Label className={ATlabelStyles.inputLabels}>Type</Label>
            <Dropdown
              selectedKey={ATFilterOptions.Type}
              placeholder="Select an option"
              options={ATDropDownOptions.Type}
              styles={ATdropdownStyles}
              onChange={(e, option: any) => {
                ATListFilter("Type", option["key"]);
              }}
            />
          </div>
          <div>
            <Label className={ATlabelStyles.inputLabels}>Area</Label>
            <Dropdown
              selectedKey={ATFilterOptions.Area}
              placeholder="Select an option"
              options={ATDropDownOptions.Area}
              styles={ATdropdownStyles}
              onChange={(e, option: any) => {
                ATListFilter("Area", option["key"]);
              }}
            />
          </div>
          <div>
            <Label className={ATlabelStyles.inputLabels}>Product</Label>
            <Dropdown
              selectedKey={ATFilterOptions.Product}
              placeholder="Select an option"
              options={ATDropDownOptions.Product}
              styles={ATdropdownStyles}
              onChange={(e, option: any) => {
                ATListFilter("Product", option["key"]);
              }}
            />
          </div>
          <div>
            <Icon
              iconName="Refresh"
              className={ATIconStyleClass.refresh}
              onClick={() => {}}
            />
          </div>
        </div>
      </div>
      <div>
        <DetailsList
          items={ATData}
          columns={_dpColumns}
          setKey="set"
          layoutMode={DetailsListLayoutMode.justified}
          selectionMode={SelectionMode.none}
          styles={gridStyles}
        />
      </div>
    </div>
  );
};

export default ActivityTemplate;
