import * as React from "react";
import { Provider, Flex, Text, Button, Header } from "@fluentui/react";
import TeamsBaseComponent, { ITeamsBaseComponentProps, ITeamsBaseComponentState } from "msteams-react-base-component";
import * as microsoftTeams from "@microsoft/teams-js";

/**
 * State for the perimeterTabDemoTabTab React component
 */
export interface IPerimeterTabDemoTabState extends ITeamsBaseComponentState {
    entityId?: string;
}

/**
 * Properties for the perimeterTabDemoTabTab React component
 */
export interface IPerimeterTabDemoTabProps extends ITeamsBaseComponentProps {

}

/**
 * Implementation of the PerimeterTabDemo content page
 */
export class PerimeterTabDemoTab extends TeamsBaseComponent<IPerimeterTabDemoTabProps, IPerimeterTabDemoTabState> {

    public componentWillMount() {
        this.updateTheme(this.getQueryVariable("theme"));
        this.setState({
            fontSize: this.pageFontSize()
        });

        if (this.inTeams()) {
            microsoftTeams.initialize();
            microsoftTeams.registerOnThemeChangeHandler(this.updateTheme);
            microsoftTeams.getContext((context) => {
                this.setState({
                    entityId: context.entityId
                });
                this.updateTheme(context.theme);
            });
        } else {
            this.setState({
                entityId: "This is not hosted in Microsoft Teams"
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
                                <Button onClick={() => alert("It worked!")}>A sample button</Button>
                            </div>
                        </div>
                    </Flex.Item>
                    <Flex.Item styles={{
                        padding: ".8rem 0 .8rem .5rem"
                    }}>
                        <Text size="smaller" content="(C) Copyright Contoso" />
                    </Flex.Item>
                </Flex>
            </Provider>
        );
    }
}
