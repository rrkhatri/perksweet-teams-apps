import * as React from "react";
import { Provider, Flex, Header, Dropdown, Input, ShorthandCollection, DropdownItemProps } from "@fluentui/react-northstar";
import { useState, useEffect, useRef } from "react";
import { useTeams } from "msteams-react-base-component";
import * as microsoftTeams from "@microsoft/teams-js";

/**
 * Implementation of the connectorConnector Connector connect page
 */
export const ConnectorConnectorConfig = () => {
    const [{ theme, context }] = useTeams();
    const [inputs, setInputs] = useState({
        email: "",
        password: ""
    });
    const handleChange = (event) => {
        const name = event.target.name;
        const value = event.target.value;
        setInputs(values => ({ ...values, [name]: value }));
    };

    useEffect(() => {
        if (context) {
            microsoftTeams.settings.registerOnSaveHandler((saveEvent: microsoftTeams.settings.SaveEvent) => {
                // INFO: Should really be of type microsoftTeams.settings.Settings, but configName does not exist in the Teams JS SDK
                const settings: any = {
                    email: inputs.email,
                    password: inputs.password,
                    contentUrl: `https://${process.env.PUBLIC_HOSTNAME}/connectorConnector/config.html?name={loginHint}&tenant={tid}&group={groupId}&theme={theme}`
                };

                microsoftTeams.settings.setSettings(settings);

                microsoftTeams.settings.getSettings((setting: any) => {
                    fetch("/api/connector/connect", {
                        method: "POST",
                        headers: [
                            ["Content-Type", "application/json"]
                        ],
                        body: JSON.stringify({
                            webhookUrl: setting.webhookUrl,
                            user: setting.userObjectId,
                            appType: setting.appType,
                            groupName: context.groupId,
                            email: settings.email,
                            password: settings.password
                        })
                    }).then(response => {
                        if (response.status === 200 || response.status === 302) {
                            saveEvent.notifySuccess();
                        } else {
                            saveEvent.notifyFailure(response.statusText);
                        }
                    }).catch(e => {
                        saveEvent.notifyFailure(e);
                    });
                });
            });

            microsoftTeams.settings.getSettings((settings: any) => {
                // setColor(availableColors.find(c => c.code === settings.entityId));
            });
        }
    }, [inputs.email, inputs.password, context]);

    useEffect(() => {
        if (context) {
            let validityState = false;
            const regex = /^(([^<>()[\].,;:\s@"]+(\.[^<>()[\].,;:\s@"]+)*)|(".+"))@(([^<>()[\].,;:\s@"]+\.)+[^<>()[\].,;:\s@"]{2,})$/i;

            if ((!inputs.email || regex.test(inputs.email)) && inputs.password) {
                validityState = true;
            }

            microsoftTeams.settings.setValidityState(validityState);
        }
    }, [inputs.email, inputs.password, context]);

    return (
        <Provider theme={theme}>
            <Flex fill={true}>
                <Flex.Item>
                    <div>
                        <h2>Configure your Connector using Perksweet Credentials</h2>

                        <Input type="email"
                            name="email"
                            value={inputs.email}
                            required={true}
                            placeholder="Email"
                            onChange={handleChange}/>

                        <br/><br/>

                        <Input type="password"
                            name="password"
                            value={inputs.password}
                            required={true}
                            placeholder="Password"
                            onChange={handleChange}/>
                    </div>
                </Flex.Item>
            </Flex>
        </Provider>
    );
};
