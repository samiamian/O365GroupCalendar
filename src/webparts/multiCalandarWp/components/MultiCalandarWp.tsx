import * as React from 'react';
import styles from './MultiCalandarWp.module.scss';
import { IMultiCalandarWpProps } from './IMultiCalandarWpProps';
import { IMultiCalandarWpState } from './IMultiCalandarWpState';
import { escape } from '@microsoft/sp-lodash-subset';
import { Calendar, momentLocalizer } from 'react-big-calendar';
import * as moment from 'moment';
import mcs from '../services/MultiCalServices';
import { Fragment } from 'react';
import "react-big-calendar/lib/css/react-big-calendar.css";


const localizer = momentLocalizer(moment);

export default class MultiCalandarWp extends React.Component<IMultiCalandarWpProps, IMultiCalandarWpState,{}> {
  constructor(props: IMultiCalandarWpProps){
    super(props);
    this.state = {
      groupID: '',
      timeZone: '',
      events: [],
    };    
  }

  public componentDidMount() {    

    //Get group ID of the Site
      mcs.getGroupGUID("sami_team_test",this.props.context)
        .then(value => {
          this.setState({ groupID: value });
        }).catch(err => {
          console.log(err);
          this.setState({ groupID: "Group ID Data not retrieved! Contact Admin."});
        });

      //get time zone of current site.
      mcs.getCurrentSiteTimeZone(this.props.context)
        .then(value => {
          this.setState({ timeZone: value});
        }).catch(err => {
          this.setState({ timeZone: "Time Zone Data not retrieved! Contact Admin."});
        });

      mcs.getAllGroupEvents("afaf4c15-0090-48ad-a4bf-e8dcbef1200c",this.props.context)
        .then(value => {
          this.setState({ events: value});
        });

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
      <section className={`${styles.multiCalandarWp} ${hasTeamsContext ? styles.teams : ''}`}>
        <div className={styles.welcome}>
          <h3>Group ID: {escape(this.state.groupID)}</h3>
          <h3>Time Zone: {escape(this.state.timeZone)}</h3>
          <Fragment>
              <div>
                <Calendar
                  localizer={localizer}
                  events={this.state.events}
                  startAccessor="start"
                  endAccessor="end"
                  popup={true}
                  style={{ height: 500 }}
                />
              </div>
          </Fragment>
        </div>
      </section>
    );
  }
}
