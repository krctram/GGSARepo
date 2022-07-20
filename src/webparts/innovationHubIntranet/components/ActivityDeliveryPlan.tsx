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
  DatePicker,
  IDatePickerStyles,
  MessageBar,
  MessageBarType,
  Spinner,
  SpinnerSize,
  TooltipHost,
  TooltipOverflowMode,
} from "@fluentui/react";
import "../ExternalRef/styleSheets/ProdStyles.css";
import { mergeStyles, mergeStyleSets } from "@fluentui/react/lib/Styling";
import styles from "./InnovationHubIntranet.module.scss";
import CustomLoader from "./CustomLoader";

const ActivityDeliveryPlan = (props: any) => {
  return (
    <div>
      <div>hi</div>
    </div>
  );
};

export default ActivityDeliveryPlan;
