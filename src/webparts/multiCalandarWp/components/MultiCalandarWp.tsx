import * as React from 'react';
import styles from './MultiCalandarWp.module.scss';
import { IMultiCalandarWpProps } from './IMultiCalandarWpProps';
import { IMultiCalandarWpState } from './IMultiCalandarWpState';
import { Calendar, momentLocalizer } from 'react-big-calendar';
import { FontIcon, Icon } from '@fluentui/react/lib/Icon';
import * as moment from 'moment';
import mcs from '../services/MultiCalServices';
import { Fragment } from 'react';
import "react-big-calendar/lib/css/react-big-calendar.css";
import { mergeStyles, Spinner, SpinnerSize, Stack } from '@fluentui/react';
import { Label } from 'office-ui-fabric-react';
import ModelEvent from "../components/ModelEvent";

const localizer = momentLocalizer(moment);

const iconClass = mergeStyles({
  fontSize: 20,
  height: 45,
  width: 30,
  float: 'right',
  cursor: 'default',
  padding: 7,
});

const icon = mergeStyles({
  alignItems: 'center',
  display: 'flex',
  height: 75,
  justifyContent: 'left',
  fontSize: 18,
  color:'red',
});

const invalidGroupStyle: React.CSSProperties = {
  alignItems: 'center',
  display: 'flex',
  height: 75,
  justifyContent: 'left',
  fontSize: 18,
  width: 900,
};

const itemRefreshStyles: React.CSSProperties = {
  alignItems: 'center',
  display: 'flex',
  height: 30,
  justifyContent: 'right',
  width: 130,
};

const itemStartStyles: React.CSSProperties = {
  alignItems: 'center',
  display: 'flex',
  height: 30,
  fontSize: 20,
  justifyContent: 'left',
};



const HomeIcon = () => <Icon iconName="ChromeClose" className={icon}/>;

export default class MultiCalandarWp extends React.Component<IMultiCalandarWpProps, IMultiCalandarWpState,{}> {


  constructor(props: IMultiCalandarWpProps){
    super(props);
    this.state = {
      groupID: '',
      isGroupIDValid: false,
      refreshed: true,
      dataLoading: false,
      calendarLoading: true,
      aEvent: null,
      isModalOpen: false,
      events: [],
    };
 
  }

  public toggleModal(){
    this.setState({isModalOpen: !this.state.isModalOpen});
  }
  public closeDialog() {  
    this.setState({ isModalOpen: false });
  } 
  public eventSelected = (event: object) => {
      this.setState({isModalOpen: true,aEvent: event});
  }

  public componentDidMount() {    
    //Get group ID of the Site
      mcs.getGroupGUID(decodeURIComponent(this.props.siteName),this.props.context)
        .then(value => {
          this.setState({ groupID: value,isGroupIDValid: true , calendarLoading: false});
          mcs.getAllGroupEvents(value,this.props.context)
          .then(value => {
            this.setState({ events: value});
          });

        }).catch(err => {
          console.log(err);
          this.setState({ groupID: "Group ID Data not retrieved! Contact Admin.", isGroupIDValid: false, calendarLoading: false});
        });
//sami_team_test
        //"afaf4c15-0090-48ad-a4bf-e8dcbef1200c"


  }

  public refreshEvents(){
    this.setState({refreshed : !this.state.refreshed, dataLoading: !this.state.dataLoading, events: []});

    mcs.getAllGroupEvents(this.state.groupID,this.props.context)
    .then(value => {
      this.setState({ events: value});
    });
    setTimeout(() => {this.setState({refreshed : !this.state.refreshed, dataLoading: !this.state.dataLoading}); }, 1000);
  }
  public render(): React.ReactElement<IMultiCalandarWpProps> {
    const {
      description,
      siteName,
      color,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    return (
      <div>
          {this.state.calendarLoading && !this.state.isGroupIDValid &&<Spinner size={SpinnerSize.large} />}
          <Stack enableScopedSelectors horizontal horizontalAlign="space-between">
            <div style={itemStartStyles}> {this.state.isGroupIDValid && <><Label>Events of {decodeURIComponent(this.props.description)}</Label></>}</div>
            {this.state.refreshed && this.state.isGroupIDValid &&<><div style={itemRefreshStyles}><FontIcon title='Refresh Event List' aria-label='Refresh' iconName='Refresh' className={iconClass} onClick={(event) => {this.refreshEvents()}}/><Label>Sync Calendar</Label></div> </>}
            {this.state.dataLoading && this.state.isGroupIDValid &&<><div style={itemRefreshStyles}><Label >Refreshing... </Label><Spinner size={SpinnerSize.large} /></div></>}
         </Stack>
        {this.state.isGroupIDValid && <>
        <div>
            <Fragment>
              <div>
                <Calendar
                  localizer={localizer}
                  events={this.state.events}
                  startAccessor="start"
                  endAccessor="end"
                  onSelectEvent={(e) => this.eventSelected(e)}
                  eventPropGetter={event => ({style: {backgroundColor: this.props.color}})}
                  style={{ height: 500 }} />
              </div>
            </Fragment>
         </div></>}
         <div className={styles.oneLine}>
            <div className={styles.square} style={{backgroundColor:this.props.color}}></div>
            <div>{this.props.siteName}</div>
        </div>
         <Stack enableScopedSelectors horizontal horizontalAlign="space-around">
            {!this.state.isGroupIDValid && !this.state.calendarLoading &&<>
                <div>
                    <HomeIcon></HomeIcon>
                </div>
                <div></div>
                <div>
                    <Label style={invalidGroupStyle}> - Unable to retrieve GroupID from the given site. Please check your site name.</Label>
                </div>
            </>}
         </Stack>
         {this.state.isModalOpen ?  <ModelEvent  
            isOpen={this.state.isModalOpen}  
            title={this.state.aEvent.title}  
            start={this.state.aEvent.start} 
            end={this.state.aEvent.end} 
            details={this.state.aEvent.bodyPreview}
            color={this.props.color}
            onClose={this.closeDialog.bind(this)}> 
          </ModelEvent> : <></>  }
      </div>
    );
  }
}