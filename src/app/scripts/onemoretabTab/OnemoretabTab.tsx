import * as React from "react";
import { Provider, Flex, Text, Button, Header } from "@fluentui/react-northstar";
import TeamsBaseComponent, { ITeamsBaseComponentState } from "msteams-react-base-component";
import * as microsoftTeams from "@microsoft/teams-js";
import * as Msal from "msal";

let res;
const msalConfig = {
    auth: {
      clientId: "3388741f-634a-4ccb-9fc5-55f5c75f1d68"
    }
};

const groupConfig = {
    "displayName": "team3kjewkj5",
    "description": "another4jndkwsdd435 random discription",
    "groupTypes": ["Unified"],
    "mailEnabled": true,
    "mailNickname": "ashu",
    "securityEnabled": false
};

const teamConfig = {
    "template@odata.bind": "https://graph.microsoft.com/v1.0/teamsTemplates('standard')",
    "visibility": "Private",
    "displayName": "Sample Engineering Team",
    "description": "This is a sample engineering team, used to showcase the range of properties supported by this API",
    "channels": [
        {
            "displayName": "Announcements",
            "isFavoriteByDefault": true,
            "description": "This is a sample announcements channel that is favorited by default. Use this channel to make important team, product, and service announcements."
        },
        {
            "displayName": "Training",
            "isFavoriteByDefault": true,
            "description": "This is a sample training channel, that is favorited by default, and contains an example of pinned website and YouTube tabs.",
            "tabs": [
                {
                    "teamsApp@odata.bind": "https://graph.microsoft.com/v1.0/appCatalogs/teamsApps('com.microsoft.teamspace.tab.web')",
                    "name": "A Pinned Website",
                    "configuration": {
                        "contentUrl": "https://docs.microsoft.com/microsoftteams/microsoft-teams"
                    }
                },
                {
                    "teamsApp@odata.bind": "https://graph.microsoft.com/v1.0/appCatalogs/teamsApps('com.microsoft.teamspace.tab.youtube')",
                    "name": "A Pinned YouTube Video",
                    "configuration": {
                        "contentUrl": "https://tabs.teams.microsoft.com/Youtube/Home/YoutubeTab?videoId=X8krAMdGvCQ",
                        "websiteUrl": "https://www.youtube.com/watch?v=X8krAMdGvCQ"
                    }
                }
            ]
        },
        {
            "displayName": "Planning",
            "description": "This is a sample of a channel that is not favorited by default, these channels will appear in the more channels overflow menu.",
            "isFavoriteByDefault": false
        },
        {
            "displayName": "Issues and Feedback",
            "description": "This is a sample of a channel that is not favorited by default, these channels will appear in the more channels overflow menu."
        }
    ],
    "memberSettings": {
        "allowCreateUpdateChannels": true,
        "allowDeleteChannels": true,
        "allowAddRemoveApps": true,
        "allowCreateUpdateRemoveTabs": true,
        "allowCreateUpdateRemoveConnectors": true
    },
    "guestSettings": {
        "allowCreateUpdateChannels": false,
        "allowDeleteChannels": false
    },
    "funSettings": {
        "allowGiphy": true,
        "giphyContentRating": "Moderate",
        "allowStickersAndMemes": true,
        "allowCustomMemes": true
    },
    "messagingSettings": {
        "allowUserEditMessages": true,
        "allowUserDeleteMessages": true,
        "allowOwnerDeleteMessages": true,
        "allowTeamMentions": true,
        "allowChannelMentions": true
    },
    "discoverySettings": {
        "showInTeamsSearchAndSuggestions": true
    },
    "installedApps": [
        {
            "teamsApp@odata.bind": "https://graph.microsoft.com/v1.0/appCatalogs/teamsApps('com.microsoft.teamspace.tab.vsts')"
        },
        {
            "teamsApp@odata.bind": "https://graph.microsoft.com/v1.0/appCatalogs/teamsApps('1542629c-01b3-4a6d-8f76-1938b779e48d')"
        }
    ]
}

// const teamConfig = {
//     "memberSettings": {
//       "allowCreatePrivateChannels": true,
//       "allowCreateUpdateChannels": true
//     },
//     "messagingSettings": {
//       "allowUserEditMessages": true,
//       "allowUserDeleteMessages": true
//     },
//     "funSettings": {
//       "allowGiphy": true,
//       "giphyContentRating": "strict"
//     },
//     "channels":[
//         {
//         "displayName":"Class Announcements 📢",
//         "isFavoriteByDefault":true
//         },
//         {
//         "displayName":"Homework 🏋️",
//         "isFavoriteByDefault":true
//         }
//     ]
//   };

const createRequest = {
    scopes: ["Group.ReadWrite.All"]
};

const loginRequest = {
    scopes: ["openid", "profile", "User.Read"]
};

// Add here scopes for access token to be used at MS Graph API endpoints.
const tokenRequest = {
    scopes: ["Mail.Read"]
};

const graphConfig = {
    graphMeEndpoint: "https://graph.microsoft.com/v1.0/me",
    graphMailEndpoint: "https://graph.microsoft.com/v1.0/me/messages",
    graphCreateGroup: "https://graph.microsoft.com/v1.0/groups",
    graphCreateTeam: "https://graph.microsoft.com/v1.0/teams"
};


const myMSALObj = new Msal.UserAgentApplication(msalConfig);

/**
 * State for the onemoretabTabTab React component
 */
export interface IOnemoretabTabState extends ITeamsBaseComponentState {
    entityId?: string;
    renderText?: string;
}

/**
 * Properties for the onemoretabTabTab React component
 */
export interface IOnemoretabTabProps {

}

/**
 * Implementation of the onemoretab Tab content page
 */
export class OnemoretabTab extends TeamsBaseComponent<IOnemoretabTabProps, IOnemoretabTabState> {

    public async componentWillMount() {
        this.updateTheme(this.getQueryVariable("theme"));


        if (await this.inTeams()) {
            microsoftTeams.initialize();
            microsoftTeams.registerOnThemeChangeHandler(this.updateTheme);
            microsoftTeams.getContext((context) => {
                microsoftTeams.appInitialization.notifySuccess();
                this.setState({
                    entityId: context.entityId,
                    renderText: "click on 'see pofile' or 'read mail' to load data"
                });
                this.updateTheme(context.theme);
            });
        } else {
            this.setState({
                entityId: "This is not hosted in Microsoft Teams",
                renderText: "click on 'see pofile' or 'read mail' to load data"
            });
        }
    }
    public signIn = () => {
        myMSALObj.loginPopup(loginRequest)
        .then(loginResponse => {
        console.log("id_token acquired at: " + new Date().toString());
        console.log(loginResponse);
        if (myMSALObj.getAccount()) {
            console.log(myMSALObj.getAccount());
        }
        }).catch(error => {
        console.log(error);
        });
    }
    public signOut = () => {
        myMSALObj.logout();
    }
    public callMSGraph = (theUrl, accToken, callback) => {
        const xmlHttp = new XMLHttpRequest();
        xmlHttp.onreadystatechange = function() {
            if (this.readyState === 4 && this.status === 200) {
                callback(this.responseText);
            }
        };
        xmlHttp.open("GET", theUrl, true);
        xmlHttp.setRequestHeader("Authorization", "Bearer " + accToken.accessToken);
        xmlHttp.send();
    }
    public getTokenPopup = (request) => {
        return myMSALObj.acquireTokenSilent(request).catch(error => {
        console.log(error);
        console.log("silent token acquisition fails. acquiring token using popup");
        // fallback to interaction when silent call fails
        return myMSALObj.acquireTokenPopup(request).then(tokenResponse => {
            return tokenResponse;
        }).catch(err1 => {
            console.log(err1);
        });
        });
    }
    public updateUI = (res1) => {
        this.setState({
            renderText: res1
        });
    }
    public seeProfile = () => {
        if (myMSALObj.getAccount()) {
            this.getTokenPopup(loginRequest)
            .then(response => {
                this.callMSGraph(graphConfig.graphMeEndpoint, response, this.updateUI);
            }).catch(err2 => {
                console.log(err2);
            });
        }
    }
    public readMail = () => {
        if (myMSALObj.getAccount()) {
            this.getTokenPopup(tokenRequest)
            .then(response => {
                this.callMSGraph(graphConfig.graphMailEndpoint, response, this.updateUI);
            }).catch(error3 => {
                console.log(error3);
            });
        }
    }
    public callGraphForGroup = (theUrl, accToken, callback) => {
        const xmlHttp = new XMLHttpRequest();
        xmlHttp.onreadystatechange = function() {
            if (this.readyState === 4 && this.status === 201) {
                res = JSON.parse(this.responseText);
                callback(res["id"]);
            }else{
                callback(this.responseText);
            }
        };
        xmlHttp.open("POST", theUrl, true);
        xmlHttp.setRequestHeader("Content-Type", "application/json");
        xmlHttp.setRequestHeader("Authorization", "Bearer " + accToken.accessToken);
        xmlHttp.send(JSON.stringify(groupConfig));
    }
    public createGroup = () => {
        if (myMSALObj.getAccount()) {
            this.getTokenPopup(createRequest)
            .then(response => {
                this.callGraphForGroup(graphConfig.graphCreateGroup, response, this.updateUI);
            }).catch(error4 => {
                console.log(error4);
            });
        }
    }
    public listGroup = (theUrl, accToken, callback) => {
        const xmlHttp = new XMLHttpRequest();
        xmlHttp.onreadystatechange = function() {
            if (this.readyState === 4 && this.status === 200) {
                res = JSON.parse(this.responseText);
                callback(theUrl, accToken, res.value[0].id, callback);
            }else{
                callback(this.responseText);
            }
        };
        xmlHttp.open("GET", theUrl, true);
        xmlHttp.setRequestHeader("Authorization", "Bearer " + accToken.accessToken);
        xmlHttp.send();
    }
    public groupToTeam = (theUrl, accToken, id, callback) => {
        const xmlHttp = new XMLHttpRequest();
        xmlHttp.onreadystatechange = function() {
            if (this.readyState === 4 && this.status === 201) {
                callback(this.responseText);
            }else{
                callback(this.responseText);
            }
        };
        xmlHttp.open("PUT", theUrl+"/"+id+"/team", true);
        xmlHttp.setRequestHeader("Content-Type", "application/json");
        xmlHttp.setRequestHeader("Authorization", "Bearer " + accToken.accessToken);
        xmlHttp.send(JSON.stringify(teamConfig));
    }
    public createTeamWithChannel = (theUrl, accToken, callback) => {
        const xmlHttp = new XMLHttpRequest();
        xmlHttp.onreadystatechange = function() {
            if (this.readyState === 4 && this.status === 201) {
                console.log(this.responseText);
                callback(this.responseText);
            }else{
                callback(this.responseText);
            }
        };
        xmlHttp.open("POST", theUrl, true);
        xmlHttp.setRequestHeader("Content-Type", "application/json");
        xmlHttp.setRequestHeader("Authorization", "Bearer " + accToken.accessToken);
        xmlHttp.send(JSON.stringify(teamConfig));
    }
    public createTeam = () => {
        if (myMSALObj.getAccount()) {
            this.getTokenPopup(createRequest)
            .then(response => {
                this.createTeamWithChannel(graphConfig.graphCreateTeam, response, this.updateUI)
                // this.listGroup(graphConfig.graphCreateGroup, response, this.groupToTeam);
            }).catch(error4 => {
                console.log(error4);
            });
        }
    }
    /**
     * The render() method to create the UI of the tab
     */
    public render() {
        return (
            <Provider theme={this.state.theme}>
                <Flex fill={true} column styles={{
                    padding: ".8rem 0 .8rem .5rem"
                }}>
                    <Flex.Item>
                        <Header content="This is your tab" />
                    </Flex.Item>
                    <Flex.Item>
                        <div>

                            <div>
                                <Text content={this.state.entityId} />
                            </div>

                            <div>
                                <Button onClick={this.signIn}>sign in</Button>
                                <Button onClick={this.signOut}>sign out</Button>
                                <Button onClick={this.seeProfile}>see profile</Button>
                                <Button onClick={this.readMail}>read mail</Button>
                                <Button onClick={this.createGroup}>create group</Button>
                                <Button onClick={this.createTeam}>create team</Button>
                            </div>

                            <div>
                                <Text content={this.state.renderText} />
                            </div>
                        </div>
                    </Flex.Item>
                    <Flex.Item styles={{
                        padding: ".8rem 0 .8rem .5rem"
                    }}>
                        <Text size="smaller" content="(C) Copyright axonfactory" />
                    </Flex.Item>
                </Flex>
            </Provider>
        );
    }
}
