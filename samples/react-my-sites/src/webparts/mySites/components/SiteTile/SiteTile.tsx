import * as React from "react";
import { getUserPhoto } from "../../../../Utils/Utils";
import { ISiteTileProps } from "./ISiteTileProps";
import { ISiteTileState } from "./ISiteTileState";
import { ITitleData } from "../../../../Entities/ITitleData";
import strings from "MySitesWebPartStrings";
import { useUserSites } from '../../../../Hooks/useUserSites';
import {
  mergeStyleSets,
  IDocumentCardStyles,
  DocumentCardTitle,
  DocumentCardPreview,
  DocumentCardDetails,
  DocumentCardActivity,
  ImageFit,
  Icon,
  IIconStyles,
  DocumentCard,
  DocumentCardType,
  IDocumentCardActivityStyles,
  ImageIcon,
} from "@fluentui/react";


const _siteLogoSP =
    "https://static2.sharepointonline.com/files/fabric-cdn-prod_20200430.002/assets/brand-icons/product/svg/sharepoint_48x1.svg";

const _siteLogoOndrive =
    "https://static2.sharepointonline.com/files/fabric-cdn-prod_20200430.002/assets/brand-icons/product/svg/onedrive_48x1.svg";

const _teamsLogo = "https://static2.sharepointonline.com/files/fabric-cdn-prod_20200430.002/assets/brand-icons/product/svg/teams_48x1.svg";

export const SiteTile: React.FunctionComponent<ISiteTileProps> = (
  props: ISiteTileProps
) => {
  const [state, setState] = React.useState<ISiteTileState>({
    hasTeam: false,
    tileData: {} as ITitleData,
  });

  const { checkGroupHasTeam } = useUserSites();
  // Global Compoment Styles
  const stylesTile = mergeStyleSets({
    imageContainer: {
      display: "flex",
      width: 104,
      height: "100%",
      alignItems: "center",
      justifyContent: "center",
      fontSize:  props.themeVariant ? props.themeVariant.fonts.superLarge.fontSize: 24,
      color: props.themeVariant ? props.themeVariant.palette.themePrimary: '',
      backgroundColor: props.themeVariant ? props.themeVariant.palette.neutralLighterAlt: '',
    },
    webPartTile: {
      paddingLeft: 30,
      paddingTop: 20,
    },
    titleContainer: {
      width: "100%",
      display: "flex",
      flexDirection: "row",
      justifyContent: "start",
    },
   imageClass: {
      marginTop: 8,
      fontSize: 20,
      height: 20,
      width: 20,
      marginRight: 7,
      color: props.themeVariant ? props.themeVariant.palette.themePrimary: '',
   }
  });

  const documentCardStyles: Partial<IDocumentCardStyles> = {
    root: {
      maxWidth: "100%",
      maxHeight: 106,
      marginTop: 10,
      marginLeft: 7,
      marginRight: 7,
    },
  };

  const groupIconStyles: Partial<IIconStyles> = {
    root: {
      fontSize: 20,
      color: props.themeVariant ? props.themeVariant.palette.themePrimary: 'white',
      marginTop: 8,
      marginRight: 7,
    },
  };
  const DocumentCardActivityStyles: Partial<IDocumentCardActivityStyles> = {
    root: { paddingBottom: 0 },
  };

  const DocumentCardDetailsStyles: Partial<IDocumentCardActivityStyles> = {
    root: { justifyContent: "flex-start" },

  };

  let _activityUserEmail = "N/A";
  let _activityUser = "N/A";
  let _activityDate = "N/A";
  let _activityMessage = "No Information";
  let _userPhoto: string = undefined;

  const {
    SiteLogo,
    Title,
    GroupId,
    LastModifiedTime,
    ModifiedById,
    SiteGroup,
    OriginalPath,
    CreatedBy,
    Created,
    IsHubSite,
    WebTemplate

  } = props.site;

  // Use Effect on Mounting
  React.useEffect(() => {
    (async () => {
      if (ModifiedById && LastModifiedTime) {
        const _modifiedBySplit = ModifiedById.split("|");
        _activityUserEmail = _modifiedBySplit[0].trim();
        _activityUser = _modifiedBySplit[1].trim();
        const _lastModified =  new Date(LastModifiedTime);
        _activityDate = _lastModified.toLocaleDateString() + ' ' + _lastModified.toLocaleTimeString();
        _activityMessage = `${strings.ChangedOnLabel}${_activityDate}`;
        try {
          if (_activityUserEmail) {
             
            _userPhoto = await getUserPhoto(_activityUserEmail);
          }
        } catch (error) {
          console.log(error);
        }
        //  _userPhoto =  `/_layouts/15/userphoto.aspx?size=M&accountname=${_activityUserEmail}`;
      } else {
        _activityUserEmail = undefined;
        _activityUser = CreatedBy;
        const _lastCreated =  new Date(Created);
        _activityDate = _lastCreated.toLocaleDateString() + ' ' + _lastCreated.toLocaleTimeString();
        
        _activityMessage = `${strings.CreatedOnLabel}${_activityDate}`;
        _userPhoto = undefined;
      }

      // If is a group check if has a team
      let _hasTeam = false;
      if (GroupId){
        _hasTeam = await checkGroupHasTeam(GroupId, props.msGraphClient);
      }
      // Update State
      setState({
        hasTeam: _hasTeam,
        tileData: {
          activityDate: _activityDate,
          activityUser: _activityUser,
          activityMessage: _activityMessage,
          activityUserEmail: _activityUserEmail,
          userPhoto: _userPhoto
        }
      });
    })();
  }, []);

  // destrectur TileData Activity
  const {
    // activityDate,
    // activityUserEmail,
    activityMessage,
    activityUser,
    userPhoto,
  } = state.tileData;

  const { hasTeam } = state;
  // Render Component
  return (
    <>
      <DocumentCard
        styles={documentCardStyles}
        type={DocumentCardType.compact}
        onClickHref={OriginalPath}
        onClickTarget={"_blank"}
      >
        {props.site.SiteLogo ? (
          <DocumentCardPreview
            className={stylesTile.imageContainer}
            previewImages={[
              {
                previewImageSrc: SiteLogo,
                width: 104,
                height: 104,
                imageFit: ImageFit.cover,
              },
            ]}
           />
        ) : (
          <DocumentCardPreview
            className={stylesTile.imageContainer}
            previewImages={[
              {
                previewImageSrc:
                  WebTemplate == "SPSPERS" ? _siteLogoOndrive : _siteLogoSP,
                width: 68,
                height: 68,
                imageFit: ImageFit.cover,
              },
            ]}
           />
        )}
        <DocumentCardDetails styles={DocumentCardDetailsStyles}>
          <div className={stylesTile.titleContainer}>
            <DocumentCardTitle title={Title} shouldTruncate styles={{root:{flexGrow:2}}} />
            {GroupId && hasTeam && (
              <ImageIcon
              title="Group has a Team"
              imageProps={{
                src: _teamsLogo,
                className: stylesTile.imageClass,
              }}
            />

            )}

            {GroupId && GroupId !== "00000000-0000-0000-0000-000000000000" && ( // (is groupId = undefined or 000000-0000-0000-0000000000000 guid) this is showned is some personal drives
              <Icon
                styles={groupIconStyles}
                iconName="Group"
                title="Office 365 Group"
               />
            )}
            {IsHubSite == "true" && (
              <Icon
                styles={groupIconStyles}
                iconName="DrillExpand"
                title="is Hub Site"
               />
            )}
           {/*  {WebTemplate == "SPSPERS" && (
              <Icon
                styles={groupIconStyles}
                iconName="onedrive"
                title="User OneDrive"
              ></Icon>
            )} */}
            {SiteGroup == "SharePoint" && !GroupId && (
              <Icon
                styles={groupIconStyles}
                iconName="SharepointAppIcon16"
                title="SharePoint Site"
               />
            )}
          </div>
          <div title={activityMessage} >
          <DocumentCardActivity
            styles={DocumentCardActivityStyles}
            activity={activityMessage}
            people={[{ name: activityUser, profileImageSrc: userPhoto }]}
          />
          </div>
        </DocumentCardDetails>
      </DocumentCard>
    </>
  );
};
