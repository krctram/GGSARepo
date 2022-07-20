import * as React from "react";
import { sp } from "@pnp/sp";
import { graph } from "@pnp/graph";
import styles from "./InnovationHubIntranet.module.scss";
import { IInnovationHubIntranetProps } from "./IInnovationHubIntranetProps";
import MainComponent from "./MainComponent";
import {
  PeoplePicker,
  PrincipalType,
} from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { Web } from "@pnp/sp/webs";

export default class InnovationHubIntranet extends React.Component<
  IInnovationHubIntranetProps,
  {}
> {
  constructor(prop: IInnovationHubIntranetProps, state: {}) {
    super(prop);
    sp.setup({
      spfxContext: this.props.context,
    });
    graph.setup({
      spfxContext: this.props.context,
    });
  }

  public render(): React.ReactElement<IInnovationHubIntranetProps> {
    return (
      <MainComponent
        context={sp}
        spcontext={this.props.context}
        graphContext={graph}
      />
    );
  }
}
