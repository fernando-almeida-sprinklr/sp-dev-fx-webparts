import { SPComponentLoader } from '@microsoft/sp-loader';

const DEFAULT_PERSONA_IMG_HASH = '7ad602295f8386b7615b582d87bcc294';
const DEFAULT_IMAGE_PLACEHOLDER_HASH = '4a48f26592f4e1498d7a478a4c48609c';
const MD5_MODULE_ID = '8494e7d7-6b99-47b2-a741-59873e42f16f';
const PROFILE_IMAGE_URL = '/_layouts/15/userphoto.aspx?size=M&accountname=';

/**
 * Gets user photo
 * @param userId
 * @returns user photo
 */
export async function getUserPhoto(userId: string) {
  const personaImgUrl = PROFILE_IMAGE_URL + userId;
  // tslint:disable-next-line: no-use-before-declare
  const url = await getImageBase64(personaImgUrl);
  // tslint:disable-next-line: no-use-before-declare
  const newHash = await getMd5HashForUrl(url);

  if (newHash !== DEFAULT_PERSONA_IMG_HASH && newHash !== DEFAULT_IMAGE_PLACEHOLDER_HASH) {
    return 'data:image/png;base64,' + url;
  } else {
    return 'undefined';
  }
}

/**
 * Get MD5Hash for the image url to verify whether user has default image or custom image
 * @param url
 */
export async function getMd5HashForUrl(url: string) {
  // tslint:disable-next-line: no-use-before-declare
  const library = await loadSPComponentById(MD5_MODULE_ID);
  try {
    if (library?.Md5Hash) {
      const convertedHash = library.Md5Hash(url);
      return convertedHash as string;
    }
  } catch (error) {
    return url;
  }
  return url;
}

/**
 * Load SPFx component by id, SPComponentLoader is used to load the SPFx components
 * @param componentId - componentId, guid of the component library
 */
export async function loadSPComponentById(componentId: string) {
  return (SPComponentLoader.loadComponentById(componentId).catch(() => undefined)) as Promise<{ Md5Hash: ((text: string) => string) } | undefined>;
}

/**
 * Gets image base64
 * @param pictureUrl
 * @returns image base64
 */
export async function getImageBase64(pictureUrl: string): Promise<string> {
  return new Promise((resolve) => {
    const image = new Image();
    image.addEventListener('load', () => {
      const tempCanvas = document.createElement('canvas');
      (tempCanvas.width = image.width), (tempCanvas.height = image.height), tempCanvas.getContext('2d').drawImage(image, 0, 0);
      let base64Str = '';
      try {
        base64Str = tempCanvas.toDataURL('image/png');
      } catch (e) {
        resolve(base64Str);
        return;
      }
      base64Str = base64Str.replace(/^data:image\/png;base64,/, '');
      resolve(base64Str);
    });
    image.src = pictureUrl;
    resolve(pictureUrl);
  });
}
