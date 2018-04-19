import * as React from 'react';
import * as ReactDOM from 'react-dom';
import{ BaseDialog, IDialogConfiguration} from '@microsoft/sp-dialog'
import {
    autobind,
    Button,
    DialogFooter,
    DialogContent
} from 'office-ui-fabric-react';

interface IDialogContentProps {
    message: string;
    close: () =>void;
}

class IDialogContent extends React.Component<IDialogContentProps,{}>{
    constructor(props){
        super(props);

    }

    public render(): JSX.Element{
        return<DialogContent
        title='Please Note'
        subText={this.props.message}
        onDismiss={this.props.close}
        showCloseButton={true}
        >
          <DialogFooter>
        <Button text='Okay' title='Okay' onClick={this.props.close} />
      </DialogFooter>
    </DialogContent>;
    }
}

export default class IOkayContent extends BaseDialog {
    public message: string;
    
    public render(): void {
        ReactDOM.render(<IDialogContent
        close={this.close}
        message={this.message}
        />, this.domElement);
    }

    public getConfig(): IDialogConfiguration {
        return {
            isBlocking: false
        };
    }
}