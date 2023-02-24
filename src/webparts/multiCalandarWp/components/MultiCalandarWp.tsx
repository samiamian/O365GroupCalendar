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


const localizer = momentLocalizer(moment);

const iconClass = mergeStyles({
  fontSize: 20,
  height: 20,
  width: 20,
  float: 'right',
  cursor: 'default',
  padding: 10,
});

const icon = mergeStyles({
  alignItems: 'center',
  display: 'flex',
  height: 75,
  justifyContent: 'left',
  fontSize: 18,
  color:'red',
});

const itemEndStyles: React.CSSProperties = {
  alignItems: 'center',
  display: 'flex',
  height: 30,
  justifyContent: 'left',
  width: 130,
};

const invalidGroupStype: React.CSSProperties = {
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
      events: [],
    };
  
  }

  public handleSelectedEvent = (event) => {
     window.alert(event.title);
  };
  public componentDidMount() {    
    //Get group ID of the Site
      mcs.getGroupGUID("sami_team_test",this.props.context)
        .then(value => {
          this.setState({ groupID: value,isGroupIDValid: true , calendarLoading: false});
        }).catch(err => {
          console.log(err);
          this.setState({ groupID: "Group ID Data not retrieved! Contact Admin.", isGroupIDValid: false, calendarLoading: false});
        });

      mcs.getAllGroupEvents("afaf4c15-0090-48ad-a4bf-e8dcbef1200c",this.props.context)
        .then(value => {
          this.setState({ events: value});
        });

  }

  public refreshEvents(){
    this.setState({refreshed : !this.state.refreshed, dataLoading: !this.state.dataLoading, events: []});

    mcs.getAllGroupEvents("afaf4c15-0090-48ad-a4bf-e8dcbef1200c",this.props.context)
    .then(value => {
      this.setState({ events: value});
    });
    setTimeout(() => {this.setState({refreshed : !this.state.refreshed, dataLoading: !this.state.dataLoading}); }, 1000);
  }
  public render(): React.ReactElement<IMultiCalandarWpProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    return (
      <div>
          {this.state.calendarLoading && !this.state.isGroupIDValid &&<Spinner size={SpinnerSize.large} />}
          <Stack enableScopedSelectors horizontal horizontalAlign="space-between">
            <div style={itemStartStyles}> 
              {this.state.isGroupIDValid && <><Label>Events of </Label></>}
            </div>
            {this.state.refreshed && this.state.isGroupIDValid &&<><div style={itemEndStyles}><FontIcon title='Refresh Event List' aria-label='Refresh' iconName='Refresh' className={iconClass} onClick={(event) => {this.refreshEvents()}}/><Label>Sync Calendar</Label></div> </>}
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
                  onSelectEvent={(event) => this.handleSelectedEvent(event)}
                  style={{ height: 500 }} />
              </div>
            </Fragment>
         </div></>}
         <Stack enableScopedSelectors horizontal horizontalAlign="space-around">
            {!this.state.isGroupIDValid && !this.state.calendarLoading &&<>
                <div>
                    <HomeIcon></HomeIcon>
                </div>
                <div></div>
                <div>
                    <Label style={invalidGroupStype}> - Unable to retrieve GroupID from the given site. Please check your site name.</Label>
                </div>
            </>}
         </Stack>
      </div>
    );
  }
}