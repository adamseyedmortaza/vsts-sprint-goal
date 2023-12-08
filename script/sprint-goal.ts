import * as RoosterJs from 'roosterjs';
import {ExtensionDataService} from "VSS/SDK/Services/ExtensionData";
import {SprintGoalApplicationInsightsWrapper} from "./SprintGoalApplicationInsightsWrapper";
import {Helpers} from "./helpers"
import * as VSSService from "VSS/Service";
import * as RestClient from "TFS/Work/RestClient";
import {TeamContext} from "TFS/Core/Contracts";
import Controls = require("VSS/Controls");
import Menus = require("VSS/Controls/Menus");
import StatusIndicator = require("VSS/Controls/StatusIndicator");
import EmojiPicker = require("vanilla-emoji-picker");
import {TeamSettingsIteration} from "TFS/Work/Contracts";

export class SprintGoal {
    // private iterationId: string;
    // private teamId: string;
    private storageUri: HTMLAnchorElement;
    private waitControl: StatusIndicator.WaitControl;
    private editor: RoosterJs.IEditor;
    private helpers: Helpers;

    private defaultConfig: SprintGoalDto = {
        sprintGoalInTabLabel: false,
        goal: "",
        details: "",
        detailsPlain: "",
        goalAchieved: false
    }

    constructor(private ai: SprintGoalApplicationInsightsWrapper) {
        try {
            this.helpers = new Helpers();

            const context = VSS.getExtensionContext();
            this.storageUri = this.getLocation(context.baseUri);

            const webContext = VSS.getWebContext();
            this.log('constructor: webContext', webContext);
            this.log('TeamId:' + webContext.team.id);
            //this.teamId = webContext.team.id;

            const config = VSS.getConfiguration();
            this.log('constructor, foregroundInstance = ' + config.foregroundInstance);

            let reloadWhenIterationChanges = false;

            if (config.foregroundInstance) { // else: config.host.background == true
                // this code runs when the form is loaded, otherwise, just load the tab

                reloadWhenIterationChanges = true;

                //this.iterationId = config.iterationId;
                this.buildWaitControl();
                this.getSettings(webContext,true).then((settings) => {
                    this.log('constructor - getSettings', settings);
                    new EmojiPicker({});
                    this.fillForm(settings);

                });

                this.buildMenuBar(webContext);


                ai.trackPageView(document.title);
            }

            // register this 'Sprint Goal' service
            VSS.register(VSS.getContribution().id, <IContributedTab>{
                pageTitle: this.getTabTitle,
                uri: undefined,
                updateContext: (ctx) => this.contextUpdated(ctx, reloadWhenIterationChanges),
                name: this.getTabTitle,
                isInvisible: (state) => false
            });
        }
        catch (e) {
            if (this.ai) this.ai.trackException(e);
        }
    }

    private getAdminPageUri = (): string => {
        const webContext = VSS.getWebContext();
        this.log('getAdminPageUri: webContext', webContext);
        const extensionId = VSS.getExtensionContext().extensionId;
        let env = ""

        if (extensionId.indexOf("-dev") >= 0) env = "-dev";
        if (extensionId.indexOf("-acc") >= 0) env = "-acc";
        return webContext.host.uri + webContext.project.name + "/_settings/keesschollaart.sprint-goal" + env + ".SprintGoalWidget.Admin";
    }

    private buildWaitControl = () => {
        const waitControlOptions: StatusIndicator.IWaitControlOptions = {
            target: $("#sprint-goal"),
            cancellable: false,
            backgroundColor: "#ffffff",
            message: "Working on your Sprint Goal..",
            showDelay: 0
        };
        this.waitControl = Controls.create(StatusIndicator.WaitControl, $("#sprint-goal"), waitControlOptions);
    }

    private contextUpdated = (ctx, reloadWhenIterationChanges: boolean) => {
        // TODO in case reloading goes stupid
        //if (ctx.iterationId == this.iterationId) return;

        if (reloadWhenIterationChanges) {
            VSS.getService(VSS.ServiceIds.Navigation).then((hostNavigationService: IHostNavigationService) => {
                //hostNavigationService.setTabTitle("my sprint goal"); // if only this was available
                hostNavigationService.reload();
            });

        }
    }

    private getLocation = (href: string): HTMLAnchorElement => {
        const l = document.createElement("a");
        l.href = href;
        return l;
    }

    private buildMenuBar = (webContext: WebContext) => {
        const menuItems: Menus.IMenuItemSpec[] = [
            { id: "save", text: "Save", icon: "icon-save" },
            { id: "settings", text: "Settings", icon: "icon-settings" }
        ];
        const menubarOptions: Menus.MenuOwnerOptions = {
            items: menuItems,
            executeAction: (args) => {
                const command = args.get_commandName();
                switch (command) {
                    case "save":
                        this.saveSettings(webContext).then(() => {
                            VSS.getService(VSS.ServiceIds.Navigation).then((navigationService: IHostNavigationService) => {
                                navigationService.reload()
                            });
                        });
                        break;
                    case "settings":
                        VSS.getService(VSS.ServiceIds.Navigation).then((navigationService: IHostNavigationService) => {
                            navigationService.navigate(this.getAdminPageUri())
                        });
                        break;
                    default:
                        alert("Unhandled action: " + command);
                        break;
                }
            }
        };

        Controls.create(Menus.MenuBar, $(".toolbar"), menubarOptions);
    }

    async getCurrentIteration(webContext: WebContext): Promise<TeamSettingsIteration> {
        try {
            //const webContext = VSS.getWebContext();
            // Constructing the TeamContext object
            const teamContext: TeamContext = {
                project: "",
                projectId: webContext.project.id,
                team: "",
                teamId: webContext.team.id
            };

            const client = VSSService.getCollectionClient(RestClient.WorkHttpClient);
            const iterations = await client.getTeamIterations(teamContext, "current");
            const currentIteration = iterations[0];

            if (currentIteration) {
                console.log("Current Iteration: ", currentIteration.name);
                return currentIteration;

            } else {
                console.log("No current iteration found");
                return null;
            }
        } catch (error) {
            console.error("Error fetching current iteration:", error);
            return null;
        }
    }


    public getTabTitle = async (): Promise<string> => {
        const webContext = VSS.getWebContext();

        this.log('getTabTitle: webContext', webContext);

        const sprintGoalCookie = await this.getSprintGoalFromCookie(webContext);

        if (!sprintGoalCookie) {
            this.log("getTabTitle: Sprint goal not yet loaded in cookie, getting it async...");
            try {
                const settings = await this.getSettings(webContext,true);
                if (settings.goal && settings.goal !== "") {
                    return "Goal: " + settings.goal;
                } else {
                    return "Goal";
                }
            } catch {
                return "Goal";
            }
        }

        if (sprintGoalCookie && sprintGoalCookie.sprintGoalInTabLabel && sprintGoalCookie.goal !== null && sprintGoalCookie.goal !== "") {
            this.log("getTabTitle: loaded title from cookie");
            return "Goal: " + sprintGoalCookie.goal;
        } else {
            this.log("getTabTitle: Cookie found but empty goal");
            return "Goal";
        }

    }


    public getSprintGoalFromCookie = async (webContext: WebContext): Promise<SprintGoalDto | undefined> => {

        //const webContext = VSS.getWebContext();
        const currentIteration = await this.getCurrentIteration(webContext);
        //return undefined;
        let sprintGoalInTabLabel = false;
        let configKey = this.helpers.getConfigKey(currentIteration.id, webContext.team.id);
        let goal = this.getCookie( configKey + "goalText");

        if (goal) {
            sprintGoalInTabLabel = (this.getCookie(configKey + "sprintGoalInTabLabel") === "true");
        } else {
            return undefined;
        }

        return {
            goal: goal,
            sprintGoalInTabLabel: sprintGoalInTabLabel,
            details: "",
            detailsPlain: "",
            goalAchieved: false
        };
    }


    public saveSettings = async (webContext: WebContext): Promise<any> => {
        this.log('saveSettings');

        if (this.waitControl) this.waitControl.startWait();

        $(".emoji-wysiwyg-editor").blur(); //ie11 hook to force WYIWYG editor to copy value to #goal input field

        const sprintConfig = <SprintGoalDto>{
            sprintGoalInTabLabel: $("#sprintGoalInTabLabelCheckbox").prop("checked") ?? this.defaultConfig.sprintGoalInTabLabel,
            goal: $("#goalInput").val() ?? this.defaultConfig.goal,
            details: this.editor.getContent() ?? this.defaultConfig.details,
            detailsPlain: this.editor.getContent(RoosterJs.GetContentMode.PlainText) ?? this.defaultConfig.detailsPlain,
            goalAchieved: $("#achievedCheckbox").prop("checked") ?? this.defaultConfig.goalAchieved
        };

        this.log('saveSettings: sprintConfig', sprintConfig);

        if (this.ai) {
            await this.ai.trackEvent("SaveSettings", <any>{
                sprintGoalInTabLabel: sprintConfig.sprintGoalInTabLabel,
                detailsUsed: `${this.editor.getContent(RoosterJs.GetContentMode.PlainText)}`.length > 10
            });
        }

        const currentIteration = await this.getCurrentIteration(webContext);

        const configIdentifierWithTeam: string = this.helpers.getConfigKey(currentIteration.id, webContext.team.id);

        this.updateSprintGoalCookie(configIdentifierWithTeam, sprintConfig);

        return VSS.getService(VSS.ServiceIds.ExtensionData)
            .then((dataService: ExtensionDataService) => {
                this.log('saveSettings: ExtensionData Service Loaded, saving for ' + configIdentifierWithTeam, sprintConfig);
                return dataService.setValue("sprintConfig." + configIdentifierWithTeam, sprintConfig);
            })
            .then((value: object) => {
                this.log('saveSettings: settings saved!', value);
                if (this.waitControl) this.waitControl.endWait();
            });
    }

    public getSettings = async (webContext: WebContext, forceReload: boolean): Promise<SprintGoalDto> => {
        this.log('getSettings');
        if (this.waitControl) this.waitControl.startWait();
        const currentGoalInCookie = await this.getSprintGoalFromCookie(webContext);

        // const webContext = VSS.getWebContext();
        const currentIteration = await this.getCurrentIteration(webContext);

        const cookieSupport = this.checkCookie();

        if (forceReload || !currentGoalInCookie || !cookieSupport) {
            const configIdentifierWithTeam = this.helpers.getConfigKey(currentIteration.id, webContext.team.id);
            const sprintGoalDto = await this.fetchSettingsFromExtensionDataService(configIdentifierWithTeam);
            this.log('getSettings bottom - configIdentifierWithTeam', configIdentifierWithTeam);
            this.log(currentIteration.id, webContext.team.id);
            this.updateSprintGoalCookie(configIdentifierWithTeam, sprintGoalDto);
            this.updateSprintGoalCookie(configIdentifierWithTeam, sprintGoalDto);
            return sprintGoalDto;
        } else {
            this.log('getSettings: fetched settings from cookie');
            return currentGoalInCookie;
        }

    }

    private fetchSettingsFromExtensionDataService = async (key: string): Promise<SprintGoalDto | null> => {
        try {
            const dataService: ExtensionDataService = await VSS.getService(VSS.ServiceIds.ExtensionData);
            this.log('getSettings: ExtensionData Service Loaded, get value by key: ' + key);

            const sprintGoalDto: SprintGoalDto = await dataService.getValue("sprintConfig." + key);
            this.log('getSettings: ExtensionData Service fetched data', sprintGoalDto);

            if (this.waitControl) this.waitControl.endWait();

            return sprintGoalDto;
        } catch (e) {
            return null;
        }
    }


    private fetchSettingsFromExtensionDataServiceOld = (key: string): IPromise<SprintGoalDto> => {
        return VSS.getService(VSS.ServiceIds.ExtensionData)
            .then((dataService: ExtensionDataService) => {
                this.log('getSettings: ExtensionData Service Loaded, get value by key: ' + key);

                try {
                    return dataService.getValue("sprintConfig." + key);
                }
                catch (e) {
                    return null;
                }
            })
            .then((sprintGoalDto: SprintGoalDto): SprintGoalDto => {
                this.log('getSettings: ExtensionData Service fetched data', sprintGoalDto);
                if (this.waitControl) this.waitControl.endWait();
                return sprintGoalDto;
            });
    }


    private updateSprintGoalCookie = (key: string, sprintGoal: SprintGoalDto) => {
        this.setCookie(key + "goalText", sprintGoal?.goal ?? this.defaultConfig.goal);
        this.setCookie(key + "sprintGoalInTabLabel", sprintGoal?.sprintGoalInTabLabel ?? this.defaultConfig.sprintGoalInTabLabel);
    }

    public fillForm = (sprintGoal: SprintGoalDto) => {
        if (!this.checkCookie()) {
            $("#cookieWarning").show();
        }

        $("#sprintGoalInTabLabelCheckbox").change(function () {
            if (this.checked) {
                $("#ditwerkniettooltip").show();
            } else {
                $("#ditwerkniettooltip").hide();
            }
        });

        const editorDiv = <HTMLDivElement>document.getElementById('detailsText');
        this.editor = RoosterJs.createEditor(editorDiv);
        if (!sprintGoal) {
            $("#sprintGoalInTabLabelCheckbox").prop("checked", false);
            $("#achievedCheckbox").prop("checked", false);
            $("#goalInput").val("");
        }
        else {
            $("#sprintGoalInTabLabelCheckbox").prop("checked", sprintGoal.sprintGoalInTabLabel);
            $("#achievedCheckbox").prop("checked", sprintGoal.goalAchieved);
            $("#goalInput").val(sprintGoal.goal);

            this.editor.setContent(sprintGoal.details);
        }
    }

    public setCookie = (key: string, value: boolean | string) => {
        const expires = new Date();
        expires.setTime(expires.getTime() + (24 * 60 * 60 * 1000));
        document.cookie = key + '=' + value + ';expires=' + expires.toUTCString() + ';domain=' + this.storageUri.hostname + ';path=/';
    }

    public getCookie(key: string) {
        const keyValue = document.cookie.match('(^|;) ?' + key + '=([^;]*)(;|$)');
        return keyValue ? keyValue[2] : null;
    }

    public checkCookie = (): boolean => {
        this.setCookie("testcookie", true);
        return (this.getCookie("testcookie") == "true");
    }

    private log = (message: string, object: any = null) => {
        if (!window.console) return;

        if (this.storageUri.href.indexOf('dev') === -1 && this.storageUri.href.indexOf('acc') === -1) return;

        if (object) {
            console.log(message, object);
            return;
        }
        console.log(message)
    }
}

export declare class EmojiPicker {
    constructor(params: any);
}

export class SprintGoalDto {
    public goal: string;
    public sprintGoalInTabLabel: boolean;
    public goalAchieved: boolean;
    public details: string;
    public detailsPlain: string;
}
