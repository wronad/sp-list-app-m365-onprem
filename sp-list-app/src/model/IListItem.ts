export interface IListItem {
  id?: number;
  courseName: string; // Title
  courseCode: string;
  courseFrequency: string;
  targetAudience: string;
}

export interface ISpListItemPayload {
  id: string;
  webUrl: string;
  siteId: string;
  fields: any;
}

export const bundleBody = (item: IListItem): string => {
  const spListItemFields = {
    fields: {
      Title: item[COURSE_NAME], // COURSE_NAME - SP list column
      COURSE_CODE: item[COURSE_CODE],
      COURSE_FREQUENCY: item[COURSE_FREQUENCY],
      TARGET_AUDIENCE: item[TARGET_AUDIENCE],
    },
  };
  return JSON.stringify(spListItemFields);
};

export const SP_SITE = "8r1bcm.sharepoint.com";

// https://entra.microsoft.com/#view/Microsoft_AAD_IAM/TenantOverview.ReactView
export const TENANT_ID = "bd04e98d-f006-4879-9451-b1096f9d1d03";

// https://8r1bcm.sharepoint.com/_layouts/15/listedit.aspx?List=%7B53eff35b-3e5a-4ef3-b87f-3baad80b982a%7D
export const LIST_ID = "53eff35b-3e5a-4ef3-b87f-3baad80b982a";

export const MS_GRAPH_URL_SP_SITE =
  "https://graph.microsoft.com/v1.0/sites/" +
  SP_SITE +
  "/lists/" +
  LIST_ID +
  "/items?expand=fields&"; // TODO $select=col-one,col-two,coln-n&

export const SP_LIST_URL =
  "https://graph.microsoft.com/v1.0/sites/" +
  SP_SITE +
  "/lists/" +
  LIST_ID +
  "/items";

// '/_layouts/15/listedit.aspx?List=%7B53eff35b-3e5a-4ef3-b87f-3baad80b982a%7D';

//'https://8r1bcm.sharepoint.com/Lists/ccUsersTrainingCourses/items';
//'8r1bcm.sharepoint.com';

export const ID = "id";
export const COURSE_NAME = "COURSE_NAME"; // Title - SP list column
export const COURSE_CODE = "COURSE_CODE";
export const COURSE_FREQUENCY = "COURSE_FREQUENCY";
export const TARGET_AUDIENCE = "TARGET_AUDIENCE";

export const COURSE_NUM = "Course No ";
export const CODE_NUM = "Course Code ";

export const extractSpListItems = (
  spListItems: ISpListItemPayload[]
): IListItem[] => {
  let items: IListItem[] = [];
  if (spListItems && spListItems.length) {
    items = spListItems.map((spItem) => {
      return {
        id: spItem.fields[ID],
        courseName: spItem.fields[COURSE_NAME],
        courseCode: spItem.fields[COURSE_CODE],
        courseFrequency: spItem.fields[COURSE_FREQUENCY],
        targetAudience: spItem.fields[TARGET_AUDIENCE],
      };
    });
  }
  return items;
};

export const mockNewListItem = (num: number): any => {
  return {
    COURSE_NAME: `${COURSE_NUM} - ${num}`,
    COURSE_CODE: `${CODE_NUM} - ${num}`,
    COURSE_FREQUENCY: "Card Holder",
    TARGET_AUDIENCE: "Initial",
  };
};
