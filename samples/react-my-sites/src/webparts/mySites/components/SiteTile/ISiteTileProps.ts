import { IReadonlyTheme } from "@microsoft/sp-component-base";
import { MSGraphClientV3 } from "@microsoft/sp-http";
// import { ISearchResult } from "@pnp/sp/search";

export interface ISiteTileProps {
  msGraphClient: MSGraphClientV3;
  site: any; /* ISearchResult;*/
  themeVariant: IReadonlyTheme | undefined;
  locale: string;
}
