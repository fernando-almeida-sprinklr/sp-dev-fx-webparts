import { MSGraphClientV3 } from "@microsoft/sp-http";
import { Filters } from "../Entities/EnumFilters";
import { sp } from "@pnp/sp";
import "@pnp/sp/sites";
import { Web } from "@pnp/sp/webs";

import { dateAdd, PnPClientStorage } from "@pnp/common";

import "@pnp/graph/groups";
import "@pnp/sp/search";
import {
  SearchResults,
  SearchQueryBuilder,
  SortDirection,
} from "@pnp/sp/search";
import { Team } from "@microsoft/microsoft-graph-types";

const storage = new PnPClientStorage();

export interface ISite {
  ParentLink: string;
  SPSiteURL: string;
  SiteID: string;
  SPWebUrl: string;
  WebId: string;
  SiteLogo: string;
  SiteClosed: string;
  RelatedHubSites: string;
  IsHubSite: string;
  GroupId: string;
  RelatedGroupId: string;
  SiteGroup: string;
  Author: string;
  CreatedBy: string;
  CreatedById: string;
  AccountName: string;
  ModifiedBy: string;
  ModifiedById: string;
  LastModifiedTime: string;
  OriginalPath: string;
  Path: string;
  Title: string;
  Created: string;
  WebTemplate: string;
}

export async function checkGroupHasTeam(
  groupId: string,
  MSGraphClientV3: MSGraphClientV3
) {
  // Check if value alreqdy cached
  const cachedValue: string = storage.local.get(groupId);
  if (!cachedValue) {
    try {
      await MSGraphClientV3.api(
        `/groups/${groupId}/team`
      ).version("V1.0")
        .get();
      // put a value into storage with an expiration
      storage.local.put(groupId, "true", dateAdd(new Date(), "day", 1));
      return true;
    } catch (error) {
      // Team don't exists or user don' have acess
      // put a value into storage with an expiration
      storage.local.put(groupId, "false", dateAdd(new Date(), "day", 1));
      return false;
    }
  } else {
    // return cached value
    return cachedValue == "true" ? true : false;
  }
}

export function getUserSites(
  searchString?: string,
  itemsPerPage?: number,
  filter?: Filters,
  site?: string
) {
  let _filter = "";
  const _searchString: string = searchString ? `Title:${searchString}*` : "";
  switch (filter) {
    case Filters.All:
      _filter = "";
      break;
    case Filters.Group:
      _filter = ` GroupId:a* OR GroupId:b* OR GroupId:c* OR GroupId:d* OR GroupId:e* OR GroupId:f* OR GroupId:g* OR GroupId:h* OR GroupId:i* OR GroupId:j* OR GroupId:k* OR GroupId:l* OR GroupId:m* OR GroupId:n* OR GroupId:o* OR GroupId:p* OR GroupId:q* OR GroupId:r* OR GroupId:s* OR GroupId:t* OR GroupId:u* OR GroupId:v* OR GroupId:w* OR GroupId:x* OR GroupId:y* OR GroupId:z* OR GroupId:1* OR GroupId:2* OR GroupId:3* OR GroupId:4* OR GroupId:5* OR GroupId:6* OR GroupId:7* OR GroupId:8* OR GroupId:9* OR GroupId:0*`;
      break;
    /*  case Filters.OneDrive:
      _filter = " WebTemplate:SPSPERS"; // OneDrive
   //   _filter = " SiteGroup:Onedrive";
      break; */
    case Filters.SharePoint:
      _filter =
        " SiteGroup:SharePoint AND NOT(GroupId:b* OR GroupId:c* OR GroupId:d* OR GroupId:e* OR GroupId:f* OR GroupId:g* OR GroupId:h* OR GroupId:i* OR GroupId:j* OR GroupId:k* OR GroupId:l* OR GroupId:m* OR GroupId:n* OR GroupId:o* OR GroupId:p* OR GroupId:q* OR GroupId:r* OR GroupId:s* OR GroupId:t* OR GroupId:u* OR GroupId:v* OR GroupId:w* OR GroupId:x* OR GroupId:y* OR GroupId:z* OR GroupId:1* OR GroupId:2* OR GroupId:3* OR GroupId:4* OR GroupId:5* OR GroupId:6* OR GroupId:7* OR GroupId:8* OR GroupId:9* OR GroupId:0*)";
      break;
    case Filters.Site:
      _filter = `Path:${site}`;

      break;
  }

  const q = SearchQueryBuilder(
    `(contentclass:STS_Site OR contentclass:STS_Web) AND -Webtemplate:SPSPERS* ${_filter} ${_searchString}`
  )
    .rowLimit(itemsPerPage || 20)
    .enableSorting.sortList({
      Property: "LastModifiedTime",
      Direction: SortDirection.Descending,
    })
    .selectProperties(
      "ParentLink",
      "SPSiteURL",
      "SiteID",
      "SPWebUrl",
      "WebId",
      "SiteLogo",
      "SiteClosed",
      "RelatedHubSites",
      "IsHubSite",
      "GroupId",
      "RelatedGroupId",
      "SiteGroup",
      "Author",
      "CreatedBy",
      "CreatedById",
      "AccountName",
      "ModifiedBy",
      "ModifiedById",
      "LastModifiedTime",
      "OriginalPath",
      "Path",
      "Title",
      "Created",
      "WebTemplate"
    );
  return sp.search(q);
}

export async function getUserWebs() {
    let searchResults: SearchResults = null;
    const q = SearchQueryBuilder(
      `(contentclass:STS_Web AND -Webtemplate:SPSPERS*)`
    )
      .rowLimit(100000)
      .selectProperties(
        "ParentLink",
        "SPSiteURL",
        "SiteID",
        "SPWebUrl",
        "WebId",
        "SiteLogo",
        "SiteClosed",
        "RelatedHubSites",
        "IsHubSite",
        "GroupId",
        "RelatedGroupId",
        "SiteGroup",
        "Author",
        "CreatedBy",
        "CreatedById",
        "AccountName",
        "ModifiedBy",
        "ModifiedById",
        "LastModifiedTime",
        "OriginalPath",
        "Path",
        "Title",
        "Created",
        "WebTemplate"
      );
    const results = await sp.search(q);
    searchResults = results; // set the current results

    return searchResults;
}

  // Get User Sites
export async function getUserGroups() {
    let searchResults: SearchResults = null;
    const _filter = ` AND GroupId:a* OR GroupId:b* OR GroupId:c* OR GroupId:d* OR GroupId:e* OR GroupId:f* OR GroupId:g* OR GroupId:h* OR GroupId:i* OR GroupId:j* OR GroupId:k* OR GroupId:l* OR GroupId:m* OR GroupId:n* OR GroupId:o* OR GroupId:p* OR GroupId:q* OR GroupId:r* OR GroupId:s* OR GroupId:t* OR GroupId:u* OR GroupId:v* OR GroupId:w* OR GroupId:x* OR GroupId:y* OR GroupId:z* OR GroupId:1* OR GroupId:2* OR GroupId:3* OR GroupId:4* OR GroupId:5* OR GroupId:6* OR GroupId:7* OR GroupId:8* OR GroupId:9* OR GroupId:0*`;
    const q = SearchQueryBuilder(
      `(contentclass:STS_Site AND -Webtemplate:SPSPERS*) ${_filter}`
    )
      .rowLimit(100000)
      .selectProperties(
        "ParentLink",
        "SPSiteURL",
        "SiteID",
        "SPWebUrl",
        "WebId",
        "SiteLogo",
        "SiteClosed",
        "RelatedHubSites",
        "IsHubSite",
        "GroupId",
        "RelatedGroupId",
        "SiteGroup",
        "Author",
        "CreatedBy",
        "CreatedById",
        "AccountName",
        "ModifiedBy",
        "ModifiedById",
        "LastModifiedTime",
        "OriginalPath",
        "Path",
        "Title",
        "Created",
        "WebTemplate"
      );
    const results = await sp.search(q);
    searchResults = results; // set the current results

    return searchResults;
}

// Get Properties for Web
export async function getSiteProperties(webUrl: string) {
  const cachedWebIdValue: { Title: string } | undefined = storage.local.get(webUrl);
  if (!cachedWebIdValue) {
    const _openWeb = Web(webUrl);
    const _webProps = await _openWeb();

    storage.local.put(webUrl, _webProps, dateAdd(new Date(), "day", 1));
    //we got all the data from the web as well
    return _webProps.Title;
  }

  // we can chain

  return cachedWebIdValue.Title;
}

export async function getUserTeams(
  userId: string,
  msGraphClient: MSGraphClientV3
) {
  const cachedListTeamsValue: Team[] = storage.local.get(userId);
  if (!cachedListTeamsValue) {
    try {
      const _listOfTeams = await msGraphClient.api(`/me/joinedTeams`)
        .version("V1.0")
        .get();
      // put a value into storage with an expiration
      storage.local.put(
        userId,
        _listOfTeams.value,
        dateAdd(new Date(), "day", 1)
      );
      return _listOfTeams.value as Team[];
    } catch (error) {
      // put a value into storage with an expiration
      storage.local.put(userId, [], dateAdd(new Date(), "day", 1));
      return [];
    }
  } else {
    // return cached value
    return cachedListTeamsValue;
  }
}


export const useUserSites = () => ({
  checkGroupHasTeam,
  getUserSites,
  getUserWebs,
  getUserGroups,
  getSiteProperties,
  getUserTeams
});
