import * as React from 'react';
import styles from './SearchPage.module.scss';
import { ISearchPageProps } from './ISearchPageProps';
import { escape } from '@microsoft/sp-lodash-subset';

import { TextField } from 'office-ui-fabric-react/lib/TextField'
import { PrimaryButton } from 'office-ui-fabric-react/lib/Button'
import { registerDefaultFontFaces } from '@uifabric/styling/lib/styles/DefaultFontStyles';
import Loading from '../../../components/Loading'
import ExecDocuments from '../../../components/ExecDocuments'

import {
  redirectToSitePage,
  siteDomain,
  getFolderUrl,
  getFiles,
  IFSObject,
  IEventListDetails, getEventListDetails,
  getFileDownloadLink,
  IExecutive,
  siteCollectionUrl,
  eventDocumentLibraryTitle

} from '../../../shared/SharePoint'

import { getQueryParameters } from '../../../shared/util'
import { PageContext } from '@microsoft/sp-page-context'
import { find } from 'lodash'
import * as moment from 'moment'
import { autobind } from 'office-ui-fabric-react/lib/Utilities';
import pnp, {SearchQueryBuilder, SearchQuery, SearchResults } from 'sp-pnp-js';
import SearchResultsList from '../../../components/SearchResultsList';
import { Breadcrumb } from 'office-ui-fabric-react/lib/Breadcrumb';

export interface ISearchPageProps {
  searchTerm: string
  searchResults: IFSObject[]
}

export interface ISearchPageState {
  searchTerm: string
  searchResults: IFSObject[]
}

export default class SearchPage extends React.Component<ISearchPageProps, ISearchPageState> {
  constructor(props: ISearchPageProps) {
    super(props)

    this.state = {
      searchTerm: "",
      searchResults: [] as IFSObject[]
    }
  }
 
  private async onSearch() {
    const searchTerm = `${this.state.searchTerm} & path:"${siteCollectionUrl}/${encodeURIComponent(eventDocumentLibraryTitle)}/*"`
    const query = {
      TrimDuplicates: false,
      EnableSorting: false,
      RowLimit: 5000,
      SelectProperties: ["Title","Author","Write","LastModifiedTime", "IsDocument","contentclass","Size","Path","Description","ServerRedirectedURL","ServerRedirectedEmbedURL","FileExtension","FileType","ParentLink","LinkingUrl","OriginalPath"],
    }
    const searchQuery = SearchQueryBuilder.create(searchTerm, query).clientType("ContentSearchRegular")

    let results = await pnp.sp.search(searchQuery)
    // console.log(results.PrimarySearchResults)
    const filteredResults = results.PrimarySearchResults
      .filter(x => x.Size > 0) //folders are Size: 0
      .map((x) =>  ({
        name: x.OriginalPath.split('/').pop(),
        type: "file",
        id: 0,
        serverRelativeUrl: x.OriginalPath,
        size: x.Size,
        modifiedBy: "",
        modified: moment(x.LastModifiedTime),
        uniqueFolderId: "",
        uniqueFileId: x.DocId.toString(),
        hasUniquePermissions: false,
        serverRedirectedEmbedUrl: x.ServerRedirectedEmbedURL,
        directAccessUsers: 0,
        isDocument: x.IsDocument,
        isContainer: x.IsContainer,
        fileExtension: x.FileExtension,
        author: x.Author
      } as IFSObject))
      
      this.setState({
        searchResults: filteredResults
      })
      // console.log(filteredResults)
  }
  @autobind
  private downloadFile(item: IFSObject) {
    const url = getFileDownloadLink(item);
    window.open(url, '_blank')
  }
  
  @autobind
  private openInBrowser(item: IFSObject) {
    const url = item.serverRedirectedEmbedUrl.split('&')[0];
    window.open(url, "_blank")
  }

  public render() {
    return (
      <div className={ styles.searchPage }>
        <div className="searchContainer">
          <div className={styles.pageBreadcrumb}>
            <Breadcrumb
              items={[
                { key: "crumb0", text: "Search", isCurrentItem: true }
              ]} />
          </div>
          <TextField
            onChanged={text => this.setState({ searchTerm: text.trim() })}
            autoFocus={true}
            className={styles.searchInput}
            onKeyUp={evt => {
              if (evt.which === 13) {
                this.onSearch()
              }
            }}
          />
          <PrimaryButton
            onClick={() => {
              this.onSearch()
            }}
          >Search</PrimaryButton>
        </div>
        <div className={styles.resultsContainer}>
          { 
            this.state.searchResults.length > 0 && (
            <SearchResultsList
              items={this.state.searchResults}
              rootFolder=""
              onDownloadFileClick={this.downloadFile}
              onOpenFileInBrowserClick={this.openInBrowser}
            />
          )}
        </div>
      </div>
    );
  }
}

