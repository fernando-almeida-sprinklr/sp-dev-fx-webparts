import * as React from "react";
import { Filters } from "../../../../Entities/EnumFilters";
import "./paginationOverride.module.scss";
import { debounce } from "lodash";
import { IMySitesProps } from "./IMySitesProps";
import {
  mergeStyleSets,
  Customizer,
  MessageBar,
  MessageBarType,
  Spinner,
  SpinnerSize,
  SearchBox,
  CommandButton,
  Stack,
  IContextualMenuProps,
  IIconProps,
  IContextualMenuItem,
  FontIcon,
  Label,
  ContextualMenuItemType,
} from "@fluentui/react";
import { WebPartTitle } from "@pnp/spfx-controls-react";
import { useUserSites } from "../../../../Hooks/useUserSites";
import { IMySitesState } from "./IMySitesState";
import { SiteTile } from "../SiteTile/SiteTile";
import { SearchResults } from "@pnp/sp/search";
import { toInteger } from "lodash";
import strings from "MySitesWebPartStrings";
import _ from "lodash";
import { Pagination } from "@mui/material";
import { MSGraphClientV3 } from '@microsoft/sp-http-msgraph';
let _searchResults: SearchResults = null;
let _msGraphClient: MSGraphClientV3 = undefined;
let _filterMenuProps: IContextualMenuProps = undefined;

// Get Hook functions
const { getUserSites, getUserWebs, getSiteProperties } = useUserSites();

export const MySites: React.FunctionComponent<IMySitesProps> = (
  props: IMySitesProps
) => {
  // Global Compoment Styles
  const stylesComponent = mergeStyleSets({
    containerTiles: {
      marginTop: 5,
      display: "grid",
      marginBottom: 10,
      gridTemplateColumns: "repeat( auto-fit, minmax(300px, 1fr) )",
      gridTemplateRows: "auto",
    },
    webPartTile: {
      fontWeight: 500,
      marginBottom: 20,
    },
  });
  // Document Card Styles

  // state
  const [state, setState] = React.useState<IMySitesState>({
    errorMessage: "",
    isLoading: true,
    sites: [],
    hasError: false,
    title: props.title,
    currentPage: 1,
    totalPages: 0,
    searchValue: "",
    currentFilter: Filters.All,
    currentFilterName: "All",
    currentSelectedSite: undefined,
    filterMenuProps: undefined,
  });

  const filterIcon: IIconProps = { iconName: "Filter" };

  // get User Sites
  const _getUserSites = async (
    searchValue?: string,
    currentFilter?: Filters,
    currentFilterName?: string,
    site?: string
  ) => {
    try {
      setState({ ...state, isLoading: true });
      const { itemsPerPage } = props;
      const searchResults = await getUserSites(
        searchValue,
        itemsPerPage,
        currentFilter,
        site
      );
      _searchResults = searchResults;
      let _totalPages = searchResults.TotalRows / itemsPerPage;
      const _modulus = searchResults.TotalRows % itemsPerPage;
      _totalPages =
        _modulus > 0 ? toInteger(_totalPages) + 1 : toInteger(_totalPages);

      setState({
        ...state,
        currentPage: 1,
        totalPages: _totalPages,
        title: props.title,
        isLoading: false,
        hasError: false,
        errorMessage: "",
        sites: _searchResults.PrimarySearchResults,
        currentFilter,
        currentSelectedSite: site,
        currentFilterName,
        // tslint:disable-next-line: no-use-before-declare
        filterMenuProps: _filterMenuProps,
      });
    } catch (error) {
      console.log(error);
      setState({
        ...state,
        hasError: true,
        isLoading: false,
        errorMessage: error.message,
      });
    }
  };

  const _FilterSites = async (filter: string, site?: string) => {

    switch (filter) {
      case "All":
        await _getUserSites(state.searchValue, Filters.All, "All");
        break;
      case "Groups":
        await _getUserSites(state.searchValue, Filters.Group, "Groups");
        break;
      /*  case "OneDrive":
         await _getUserSites(state.searchValue, Filters.OneDrive, "OneDrive");
         break; */
      case "SharePoint":
        await _getUserSites(state.searchValue, Filters.SharePoint, "SharePoint");
        break;
      default:
        await _getUserSites(state.searchValue, Filters.Site, filter, site);
    }
  };


  // useEffect component did mount or modified
  React.useEffect(() => {
    (async () => {
      _msGraphClient = await props.context.msGraphClientFactory.getClient("3");

      const _sitesWithSubSites = await getUserWebs();

      const _uniqweb = _.uniqBy(
        _sitesWithSubSites.PrimarySearchResults,
        "ParentLink"
      );

      const { enableFilterO365groups, enableFilterSharepointSites, enableFilterSitesWithSubWebs } = props;

      _filterMenuProps = {
        items: [
          {
            key: "0",
            text: "All",
            iconProps: { iconName: "ThumbnailView" },
            onClick: (
              ev:
                | React.MouseEvent<HTMLElement, MouseEvent>
                | React.KeyboardEvent<HTMLElement>,
              item: IContextualMenuItem
            ) => {
              _FilterSites(item.text);
            },
          },
        ],
      };


      if (enableFilterSharepointSites) {

        _filterMenuProps.items.push({
          key: "1",
          text: "SharePoint",
          iconProps: { iconName: "SharepointAppIcon16" },
          onClick: (
            ev:
              | React.MouseEvent<HTMLElement, MouseEvent>
              | React.KeyboardEvent<HTMLElement>,
            item: IContextualMenuItem
          ) => {
            _FilterSites(item.text);
          }
        }
        );
      }

      if (enableFilterO365groups) {
        _filterMenuProps.items.push({
          key: "2",
          text: "Groups",
          iconProps: { iconName: "Group" },
          onClick: (
            ev:
              | React.MouseEvent<HTMLElement, MouseEvent>
              | React.KeyboardEvent<HTMLElement>,
            item: IContextualMenuItem
          ) => {
            _FilterSites(item.text);
          }
        }
        );
      }

      if (enableFilterSitesWithSubWebs && _sitesWithSubSites.PrimarySearchResults.length > 0) {
        _filterMenuProps.items.push(
          {
            key: 'sites',
            itemType: ContextualMenuItemType.Header,
            text: 'Sites with sub sites',
            itemProps: {
              lang: 'en-us',
            },
          }
        );

        // Add Site Collections with sub
        for (const web of _uniqweb) {
          const _webTitle = await getSiteProperties(web.ParentLink);

          // tslint:disable-next-line: no-use-before-declare
          _filterMenuProps.items.push({
            key: web.ParentLink,
            text: _webTitle,
            iconProps: { iconName: "DrillExpand" },
            onClick: (
              _:
                | React.MouseEvent<HTMLElement, MouseEvent>
                | React.KeyboardEvent<HTMLElement>,
              item: IContextualMenuItem
            ) => {
              // tslint:disable-next-line: no-use-before-declare
              _FilterSites(item.text, item.key);
            },
          });
        }
      }

      await _getUserSites("", state.currentFilter, state.currentFilterName);
    })();
  }, [props]);

  // On Search Sites
  const _onSearch = async (value: string) => {
    await _getUserSites(value, state.currentFilter, state.currentFilterName, state.currentSelectedSite);
  };

  const searchWithDebounce = debounce(_onSearch, props.searchSettings.debounceDelayMs);

  const _onChange = props.searchSettings?.debounce
    ? (_: React.ChangeEvent<HTMLInputElement>, newValue?: string) => {
      if (newValue?.length < props.searchSettings.debounceMinChars) {
        return;
      }
      // eslint-disable-next-line no-debugger
      searchWithDebounce.cancel();
      searchWithDebounce(newValue);
    }
    : undefined;

  // On Search Sites
  const _onClear = async (searchValue?: string) => {
    await _getUserSites(searchValue || '', state.currentFilter, state.currentFilterName, state.currentSelectedSite);
  };

  // Render component
  if (state.hasError) {
    // render message error
    return (
      <MessageBar messageBarType={MessageBarType.error}>
        {state.errorMessage}
      </MessageBar>
    );
  }

  // Render list of tiles
  return (
    <>
      <Customizer settings={{ theme: props.themeVariant }}>
        <WebPartTitle
          displayMode={props.displayMode}
          title={state.title}
          themeVariant={props.themeVariant}
          updateProperty={props.updateProperty}
          className={stylesComponent.webPartTile}
        />
        <Stack horizontal verticalAlign="center" horizontalAlign='start' wrap tokens={{ childrenGap: 5 }} styles={{ root: { width: '100%' } }}>
          <Stack.Item grow align="stretch">
            <SearchBox
              placeholder="Search my sites"
              styles={{ root: { width: '100%', marginBottom: 10 } }}
              onChange={_onChange}
              onSearch={_onSearch}
              onClear={() => _onClear()}
            />
          </Stack.Item>
          <Stack.Item>
            <CommandButton
              iconProps={{ iconName: "refresh" }}
              onClick={() => _onClear(state.searchValue)}
              title={strings.RefreshLabel}
            />
            <CommandButton
              iconProps={filterIcon}
              text={state.currentFilterName}
              menuProps={state.filterMenuProps}
              disabled={false}
              checked={true}
              title="filter"
            />
          </Stack.Item>
        </Stack>
        {state.isLoading ? (
          <Spinner
            size={SpinnerSize.medium}
            label={strings.LoadingLabel}
          />
        ) : (
          <>
            {
              // has sites ?
              state.sites.length > 0 ? (
                <div className={stylesComponent.containerTiles}>
                  {state.sites.map((site) => {
                    return (
                      <SiteTile
                        key={site.id}
                        site={site}
                        msGraphClient={_msGraphClient}
                        themeVariant={props.themeVariant}
                        locale={props.context.pageContext.cultureInfo.currentCultureName}
                      />
                    );
                  })}
                </div>
              ) : (
                <>
                  <Stack
                    horizontal
                    verticalAlign="center"
                    horizontalAlign="center"
                    tokens={{ childrenGap: 20 }}
                    styles={{ root: { marginTop: 50 } }}
                  >
                    <FontIcon
                      iconName="Tiles"
                      style={{ fontSize: 48 }}
                    />
                    <Label styles={{ root: { fontSize: 26 } }}>
                      No Sites Found{" "}
                    </Label>
                  </Stack>
                </>
              )
            }

            {state.totalPages > 1 && (
              <>
                <div
                  style={{
                    display: "flex",
                    justifyContent: "center",
                    marginTop: 30,
                  }}
                >
                  <Pagination
                    color="primary"
                    count={state.totalPages}
                    page={state.currentPage}
                    size="small"
                    siblingCount={0}
                    onChange={async (_, page: number) => {
                      const rs = await _searchResults.getPage(page);
                      _searchResults = rs;
                      setState({
                        ...state,
                        currentPage: page,
                        sites: _searchResults.PrimarySearchResults,
                      });
                    }}
                  />
                </div>
              </>
            )}
          </>
        )}
      </Customizer>
    </>
  );
};
