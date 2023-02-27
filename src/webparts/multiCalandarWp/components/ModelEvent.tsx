import { IModelEventProps } from './IModelEventProps';

import * as React from 'react';
import { initializeIcons } from '@fluentui/react/lib/Icons';
import styles from './ModalEvent.module.scss';

import { DefaultButton, Dialog, DialogFooter, DialogType, IDialogStyles } from '@fluentui/react';
initializeIcons();

const dialogStyles = { 
  main: { maxWidth: 900 } ,
};  
const custStyles = props => ({
  root: [{background: props.theme.palette.themePrimary,}]
});
const dialogContentProps = {  
  type: DialogType.normal,  
  title: 'Event Details',  
 //className: styles.title,
}; 

const modalProps = {  
  isBlocking: true,  
} 


export default class ModelEvent extends React.Component<IModelEventProps, {}> {

  constructor(props: IModelEventProps){
    super(props);
    this.handleClose = this.handleClose.bind(this);
  }

  public handleClose() {  
    this.props.onClose();
  }; 



  public render(): React.ReactElement<IModelEventProps> {
    const buttonStyles = {
      root: { backgroundColor: this.props.color, color: 'white' }
    };
    return (
      <div>
        <Dialog  
          hidden={!this.props.isOpen}  
          dialogContentProps={dialogContentProps}  
          styles={dialogStyles}  
          modalProps={modalProps}
          onDismiss={this.handleClose}
          maxWidth={900}
          >
            <div className={styles.detailsGrid}>  
              <div><strong>Title</strong></div>  
              <div>{this.props.title}</div>  
              <div><strong>Start</strong></div>  
              <div>{this.props.start.toString()}</div>  
              <div><strong>End</strong></div>  
              <div>{this.props.end.toString()}</div> 
              <div><strong>Detail</strong></div>  
              <div>{this.props.details}</div> 
            </div>
            <DialogFooter>  
              <DefaultButton onClick={this.handleClose} text="Cancel" styles={buttonStyles} />  
            </DialogFooter> 
        </Dialog>  
      </div>
    );
  }
}


