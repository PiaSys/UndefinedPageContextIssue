import * as React from 'react';
import styles from './ShowPageContextIssue.module.scss';
import { IShowPageContextIssueProps } from './IShowPageContextIssueProps';
import { IShowPageContextIssueState } from './IShowPageContextIssueState';
import { escape } from '@microsoft/sp-lodash-subset';

import { WebPartTitle } from "@pnp/spfx-controls-react/lib/WebPartTitle";
import { Spinner, SpinnerSize } from 'office-ui-fabric-react';
import { GridLayout } from "@pnp/spfx-controls-react/lib/GridLayout";
import {
  DocumentCard,
  DocumentCardActivity,
  DocumentCardPreview,
  DocumentCardDetails,
  DocumentCardTitle,
  IDocumentCardPreviewProps,
  DocumentCardLocation,
  DocumentCardType
} from 'office-ui-fabric-react/lib/DocumentCard';
import { ImageFit } from 'office-ui-fabric-react/lib/Image';
import { ISize } from 'office-ui-fabric-react/lib/Utilities';

// Import PnPjs supporting types
import { SPFI } from '@pnp/sp';
import { getSP } from '../pnpjsConfig';
import { IList } from '@pnp/sp/lists';
import { IContentTypes } from '@pnp/sp/content-types';
import '@pnp/sp/content-types/list';
import * as _ from 'lodash';

import { IDataItemsService } from '../../../services/itemsService/IDataItemsService';
import { DataItemsService } from '../../../services/itemsService/DataItemsService';
import { IDataItem } from '../../../services/itemsService/IDataItem';

const fakeItems: IDataItem[] = [{
  id: 0,
  title: "Adventures in SPFx",
  name: "Perry Losselyong",
  profileImageSrc: "https://robohash.org/blanditiisadlabore.png?size=50x50&set=set1",
  location: "SharePoint",
  activity: "3/13/2019"
}, {
  id: 1,
  title: "The Wild, Untold Story of SharePoint!",
  name: "Ebonee Gallyhaock",
  profileImageSrc: "https://robohash.org/delectusetcorporis.bmp?size=50x50&set=set1",
  location: "SharePoint",
  activity: "6/29/2019"
}, {
  id: 2,
  title: "Low Code Solutions: PowerApps",
  name: "Seward Keith",
  profileImageSrc: "https://robohash.org/asperioresautquasi.jpg?size=50x50&set=set1",
  location: "PowerApps",
  activity: "12/31/2018"
}, {
  id: 3,
  title: "Not Your Grandpa's SharePoint",
  name: "Sharona Selkirk",
  profileImageSrc: "https://robohash.org/velnammolestiae.png?size=50x50&set=set1",
  location: "SharePoint",
  activity: "11/20/2018"
}, {
  id: 4,
  title: "Get with the Flow",
  name: "Boyce Batstone",
  profileImageSrc: "https://robohash.org/nulladistinctiomollitia.jpg?size=50x50&set=set1",
  location: "Flow",
  activity: "5/26/2019"
}];

export default class ShowPageContextIssue extends React.Component<IShowPageContextIssueProps, IShowPageContextIssueState> {

  private _dataItemsService: IDataItemsService = null;
  private _sp: SPFI = null;

  constructor(props: IShowPageContextIssueProps) {
    super(props);

    this._sp = getSP(this.props.context);

    this.state = {
      items: props.useFakeData ? fakeItems : null,
      loading: false
    };
  }

  public async componentDidMount() {

    console.log('componentDidMount');
    console.log(this.props.context);
    console.log(this.props.context.pageContext);

    // If we don't have source list, we simply exit
    if (this.props.sourceList == null && !this.props.useFakeData) {
      return;
    }

    this._dataItemsService = new DataItemsService(this._sp, this.props.aadHttpClient, this.props.sourceList);

    await this._loadItems();    
  }

  public async componentDidUpdate(prevProps: IShowPageContextIssueProps) {

    console.log('componentDidUpdate');
    console.log(this.props.context);
    console.log(this.props.context.pageContext);

    // Refresh data if and only if query criteria changed
    if (prevProps.useFakeData === this.props.useFakeData &&
        prevProps.sourceList === this.props.sourceList) {
        return;
    } 
    
    // If we don't have source list, we simply exit
    if (this.props.sourceList == null && !this.props.useFakeData) {
      return;
    }
    
    // Refresh Assets Service instance, if switch about Mock Data changed
    if (prevProps.useFakeData !== this.props.useFakeData ||
      prevProps.sourceList !== this.props.sourceList) {

        console.log('componentDidUpdate:RefreshingData');
        console.log(this.props.context);
        console.log(this.props.context.pageContext);
    
        this._dataItemsService = new DataItemsService(this._sp, this.props.aadHttpClient, this.props.sourceList);

        await this._loadItems();    
    }
  }

  private _loadItems = async (): Promise<void> => {

    // Set loading state
    this.setState({
      loading: true
    });
    
    // Make the actual query for items
    const items = this.props.useFakeData ? fakeItems : await this._dataItemsService.GetItems();

    // Bind data
    this.setState({
      loading: false,
      items: items
    });
  }

  public render(): React.ReactElement<IShowPageContextIssueProps> {
    const {
      isDarkTheme,
      hasTeamsContext
    } = this.props;

    const { 
      items 
    } = this.state;

    return (
      <section className={`${styles.showPageContextIssue} ${hasTeamsContext ? styles.teams : ''}`}>
        <div>
          <WebPartTitle displayMode={this.props.displayMode}
                title={this.props.title}
                updateProperty={this.props.updateProperty}
                />
        </div>
        <div className={styles.welcome}>
          <img alt="" src={isDarkTheme ? require('../assets/welcome-dark.png') : require('../assets/welcome-light.png')} className={styles.welcomeImage} />
        </div>
        { this.state.loading  ?
          <Spinner size={SpinnerSize.large} /> 
          : null
        }
        { !this.state.loading && this.state.items != null && this.state.items.length > 0 ?        
          <div>
            <h3>Here are the items from the source list!</h3>
            <div>
              <GridLayout
                items={items}
                onRenderGridItem={this._onRenderGridItem}
              />
            </div>
          </div> 
          : null 
        }
      </section>
    );
  }

  private _onRenderGridItem = (item: any, finalSize: ISize, isCompact: boolean): JSX.Element => {

    return <div
      data-is-focusable={true}
      role="listitem"
      aria-label={item.title}
    >
      <DocumentCard
        type={isCompact ? DocumentCardType.compact : DocumentCardType.normal}
        onClick={(ev: React.SyntheticEvent<HTMLElement>) => alert("You clicked on a grid item")}

      >
        {!isCompact && <DocumentCardLocation location={item.location} />}
        <DocumentCardDetails>
          <DocumentCardTitle
            title={item.title}
            shouldTruncate={true}
          />
          <DocumentCardActivity
            activity={item.activity}
            people={[{ name: item.name, profileImageSrc: item.profileImageSrc }]}
          />
        </DocumentCardDetails>
      </DocumentCard>
    </div>;
  }
}
