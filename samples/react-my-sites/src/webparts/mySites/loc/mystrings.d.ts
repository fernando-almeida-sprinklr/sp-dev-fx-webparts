declare interface IMySitesWebPartStrings {
  CreatedOnLabel: string;
  ChangedOnLabel: string;
  RefreshLabel: string;
  LoadingLabel: string;
  PropertyPaneDescription: string;
  BasicGroupName: string;
  TitleFieldLabel: string;
  ItemsPerPageLabel: string;
  DebounceSearches: string;
  SearchDebounceMinChars: string;
  SearchDebounceDelayMs: string;
}

declare module 'MySitesWebPartStrings' {
  const strings: IMySitesWebPartStrings;
  export = strings;
}
