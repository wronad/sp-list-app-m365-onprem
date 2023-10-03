const ID = "id";
const COURSE_NAME = "COURSE_NAME"; // Title - SP list column
const COURSE_CODE = "COURSE_CODE";
const COURSE_FREQUENCY = "COURSE_FREQUENCY";
const TARGET_AUDIENCE = "TARGET_AUDIENCE";

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

export const bundleBodyForOnline = (item: IListItem): string => {
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

export const bundleDataForOnPrem = (
  item: IListItem,
  itemType: string
): string => {
  const data = {
    __metadata: {
      type: itemType,
    },
    Title: item[COURSE_NAME], // COURSE_NAME - SP list column
    COURSE_CODE: item[COURSE_CODE],
    COURSE_FREQUENCY: item[COURSE_FREQUENCY],
    TARGET_AUDIENCE: item[TARGET_AUDIENCE],
  };
  return JSON.stringify(data);
};

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
  const mock = {};
  mock[COURSE_NAME] = `Course No - ${num}`;
  mock[COURSE_CODE] = `Course Code - ${num}`;
  mock[COURSE_FREQUENCY] = "Card Holder";
  mock[TARGET_AUDIENCE] = "Initial";
  return mock;
};
