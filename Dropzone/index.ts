import { IInputs, IOutputs } from "./generated/ManifestTypes";
import { Landing } from "./Landing/Landing";
import * as React from "react";
import { CONTROL_VERSION } from "./version";

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
    console.log(`Modern Dropzone PCF ${CONTROL_VERSION} initialised`);
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
