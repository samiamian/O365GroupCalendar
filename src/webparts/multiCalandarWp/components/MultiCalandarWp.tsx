import * as React from 'react';
import styles from './MultiCalandarWp.module.scss';
import { IMultiCalandarWpProps } from './IMultiCalandarWpProps';
import { IMultiCalandarWpState } from './IMultiCalandarWpState';
import { escape } from '@microsoft/sp-lodash-subset';
import { Calendar, momentLocalizer } from 'react-big-calendar';
import { FontIcon } from '@fluentui/react/lib/Icon';
import * as moment from 'moment';
import mcs from '../services/MultiCalServices';
import { Fragment } from 'react';
import "react-big-calendar/lib/css/react-big-calendar.css";
import { ILabelStyles, mergeStyles, Stack } from '@fluentui/react';
import { DefaultPalette, registerIcons } from '@fluentui/react/lib/Styling';
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

const itemEndStyles: React.CSSProperties = {
  alignItems: 'center',
  display: 'flex',
  height: 30,
  justifyContent: 'left',
  width: 130,
};

const itemStartStyles: React.CSSProperties = {
  alignItems: 'center',
  display: 'flex',
  height: 30,
  justifyContent: 'left',
}
export default class MultiCalandarWp extends React.Component<IMultiCalandarWpProps, IMultiCalandarWpState,{}> {
  constructor(props: IMultiCalandarWpProps){
    super(props);
    this.state = {
      groupID: '',
      refreshing: false,
      loaded: true,
      events: [],
    };    
  }

  public componentDidMount() {    

    console.log("component did mount");
    //Get group ID of the Site
      mcs.getGroupGUID("sami_team_test",this.props.context)
        .then(value => {
          this.setState({ groupID: value });
        }).catch(err => {
          console.log(err);
          this.setState({ groupID: "Group ID Data not retrieved! Contact Admin."});
        });

      mcs.getAllGroupEvents("afaf4c15-0090-48ad-a4bf-e8dcbef1200c",this.props.context)
        .then(value => {
          this.setState({ events: value});
        });

  }

  public refreshEvents(){
    this.setState({refreshing: true,loaded: false, events: []});
    mcs.getAllGroupEvents("afaf4c15-0090-48ad-a4bf-e8dcbef1200c",this.props.context)
    .then(value => {
      this.setState({ events: value});
    });
    this.setState({refreshing: false,loaded: true});

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
          <Stack enableScopedSelectors horizontal horizontalAlign="space-between">
            <div style={itemStartStyles}> 
              <Label>Events of </Label>
            </div>
            <div style={itemEndStyles}>
              <FontIcon title='Refresh Event List' aria-label='Refresh' iconName='Refresh' className={iconClass} onClick={(event) => {this.refreshEvents()}}/> 
              <Label >Sync Calendar</Label>
            </div>
         </Stack>
        <div>
            <Fragment>
              <div>
                <Calendar
                  localizer={localizer}
                  events={this.state.events}
                  startAccessor="start"
                  endAccessor="end"
                  popup={true}
                  style={{ height: 500 }} />
              </div>
            </Fragment>
         </div>
      </div>
    );
  }
}
