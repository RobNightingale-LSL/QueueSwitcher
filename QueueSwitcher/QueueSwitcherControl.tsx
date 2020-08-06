import * as React from "react";
import { ChoiceGroup, IChoiceGroupOption } from 'office-ui-fabric-react/lib/ChoiceGroup';
import { initializeIcons } from 'office-ui-fabric-react';

initializeIcons();

export interface IQueueSwitcherControlState {
    Queues: IChoiceGroupOption[];
    SelectedQueue?: string;
}

export interface IQueueSwitcherControlProps extends IQueueSwitcherControlState {
    OnQueueChanged: (queueId: string) => void;
}

export class QueueSwitcherControl extends React.Component<IQueueSwitcherControlProps, IQueueSwitcherControlState> {
    constructor(props: IQueueSwitcherControlProps) {
        super(props);

        this.state = {
            Queues: props.Queues,
            SelectedQueue: props.SelectedQueue
        };
    }

    render() {
        return <ChoiceGroup 
            selectedKey={this.state.SelectedQueue} 
            options={this.state.Queues} 
            onChange={(ev?: React.FormEvent<HTMLElement | HTMLInputElement>, option?: IChoiceGroupOption) => {
                if (option) {
                    this.setState({
                        SelectedQueue: option.key
                    }, () => {
                        this.props.OnQueueChanged(option.key)
                    });
                }
            }}/>;
    }
}
