import { ExtensionDataService } from "VSS/SDK/Services/ExtensionData";
import * as VSS_Service from "VSS/Service";
import * as Tfs_Core_WebApi from "TFS/Core/RestClient";
import * as Tfs_Work_WebApi from "TFS/Work/RestClient";
import { TcmHttpClient } from "TFS/TestManagement/VSS.Tcm.WebApi";

import * as contract from "TFS/Core/Contracts"
import { Helpers } from "./helpers";
import { SprintGoalDto } from "./sprint-goal";
import { ExportEntry } from "./ExportEntry";

export class SprintGoalAdmin {


    private form: HTMLFormElement = <HTMLFormElement>document.getElementById('payment-form');
    private telemetryCheckbox = <HTMLInputElement>document.getElementById("telemetryCheckbox");
    private exportButton = <HTMLButtonElement>document.getElementById("exportButton");
    private helpers: Helpers;


    constructor() {
        this.helpers = new Helpers();
    }

    public load = async (): Promise<void> => {
        this.telemetryCheckbox.checked = await this.getTelemetryOptOut();
        this.telemetryCheckbox.onchange = (e) => this.setTelemetryOptOut(this.telemetryCheckbox.checked);
        this.form.addEventListener('submit', this.onFormSubmit);
        this.exportButton.onclick = (e) => this.exportButtonClick();
    }

    private onFormSubmit = async (event): Promise<void> => {
        event.preventDefault();
    }

    private getTelemetryOptOut = async (): Promise<boolean> => {
        const dataService = <ExtensionDataService>await VSS.getService(VSS.ServiceIds.ExtensionData);
        let telemetryOptOut = false;
        try {
            telemetryOptOut = await dataService.getValue<boolean>("telemetryOptOut");
        }
        catch{
            // swallow
        }
        return telemetryOptOut;
    }

    private setTelemetryOptOut = async (value: boolean): Promise<void> => {
        const dataService = <ExtensionDataService>await VSS.getService(VSS.ServiceIds.ExtensionData);
        await dataService.setValue("telemetryOptOut", !!value);
    }

    private exportButtonClick = async (): Promise<void> => {

        this.exportButton.disabled = true;
        this.exportButton.innerText = "Generating file export...";

        const dataService = <ExtensionDataService>await VSS.getService(VSS.ServiceIds.ExtensionData);
        const project = VSS.getWebContext().project;
        const result: ExportEntry[] = [];

        const workApi = Tfs_Work_WebApi.getClient();
        const collectionClient = VSS_Service.getCollectionClient(Tfs_Core_WebApi.CoreHttpClient4);
        const teams = await collectionClient.getTeams(project.id);

        const keysToDownload: {
            [key: string]: {
                teamId: string,
                teamName: string,
                iterationId: string,
                iterationName: string
            }
        } = {};
        for (let j = 0; j < teams.length; j++) {
            let team = teams[j];
            const teamContext: contract.TeamContext = {
                projectId: project.id,
                teamId: team.id,
                project: "",
                team: ""
            };

            const iterations = await workApi.getTeamIterations(teamContext);
            for (let i = 0; i < iterations.length; i++) {
                let iteration = iterations[i];
                const configKey = this.helpers.getConfigKey(iteration.id, team.id);
                keysToDownload["sprintConfig." + configKey] = {
                    teamId: team.id,
                    teamName: team.name,
                    iterationName: iteration.name,
                    iterationId: iteration.id
                };
            }
        }

        try {
            const goals = <{ [key: string]: SprintGoalDto }>await dataService.getValues(Object.keys(keysToDownload));
            for (let key in goals) {
                result.push({
                    details: goals[key].details,
                    detailsPlain: goals[key].detailsPlain,
                    goal: goals[key].goal,
                    goalAchieved: goals[key].goalAchieved,
                    sprintGoalInTabLabel: goals[key].sprintGoalInTabLabel,
                    iterationId: keysToDownload[key].iterationId,
                    iterationName: keysToDownload[key].iterationName,
                    teamId: keysToDownload[key].teamId,
                    teamName: keysToDownload[key].teamName,
                    projectId: project.id,
                    projectName: project.name
                });
            }

        }
        catch (e) {
            this.exportButton.innerText = e;
        }
        this.download("goals.json", JSON.stringify(result));

        this.exportButton.disabled = false;
        this.exportButton.innerText = "Export";
    }

    private download = (filename, text) => {
        const element = document.createElement('a');
        element.setAttribute('href', 'data:text/plain;charset=utf-8,' + encodeURIComponent(text));
        element.setAttribute('download', filename);

        element.style.display = 'none';
        document.body.appendChild(element);

        element.click();

        document.body.removeChild(element);
    }
}

VSS.ready(function () {
    VSS.require([], () => {
        const licenseAdmin = new SprintGoalAdmin();
        licenseAdmin.load().then(() => {
            VSS.register(VSS.getContribution().id, licenseAdmin);
            VSS.notifyLoadSucceeded();
        });
    });
});

