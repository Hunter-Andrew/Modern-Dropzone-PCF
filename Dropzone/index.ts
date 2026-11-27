import { IInputs, IOutputs } from "./generated/ManifestTypes";
import { Landing } from "./Landing/Landing";
import * as React from "react";
import { getLocalString } from "./utils";
import { LocalStrings } from "./consts/LocalStrings";

export class Dropzone
  implements ComponentFramework.ReactControl<IInputs, IOutputs>
{
  private theContainer: HTMLDivElement;
  private notifyOutputChanged: () => void;
  private webAPI: ComponentFramework.WebApi;
  public init(
    context: ComponentFramework.Context<IInputs>,
    notifyOutputChanged: () => void,
    state: ComponentFramework.Dictionary,
    container: HTMLDivElement
  ): void {
    const appName = getLocalString(context, LocalStrings.Button.Label_Preview)
    console.log(`${appName} PCF 2.9.5 Initialised`);
    this.theContainer = container;
    this.notifyOutputChanged = notifyOutputChanged;
    this.webAPI = context.webAPI;
  }

  public updateView(
    context: ComponentFramework.Context<IInputs>
  ): React.ReactElement {
    const isDisabled = context.mode.isControlDisabled;
    return React.createElement(Landing, { context: context, isDisabled });
  }
  public getOutputs(): IOutputs {
    return {};
  }
  public destroy(): void {}
}
