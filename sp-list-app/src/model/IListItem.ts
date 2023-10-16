const ID = "Id";
const COURSE_NAME = "COURSE_NAME"; // Title - SP list column
const COURSE_CODE = "COURSE_CODE";
const COURSE_FREQUENCY = "COURSE_FREQUENCY";
const TARGET_AUDIENCE = "TARGET_AUDIENCE";

export const LIST_FIELDS = `${ID}, Title, ${COURSE_CODE}, ${COURSE_FREQUENCY}, ${TARGET_AUDIENCE}`;

export interface IListItem {
  id?: number;
  courseName: string; // Title
  courseCode: string;
  courseFrequency: string;
  targetAudience: string;
}

interface ISpListItem {
  id?: number;
  Title: string; // COURSE_NAME
  COURSE_CODE: string;
  COURSE_FREQUENCY: string;
  TARGET_AUDIENCE: string;
}

export const bundleItem = (item: IListItem): ISpListItem => {
  return {
    Title: item[COURSE_NAME], // COURSE_NAME - SP list column
    COURSE_CODE: item[COURSE_CODE],
    COURSE_FREQUENCY: item[COURSE_FREQUENCY],
    TARGET_AUDIENCE: item[TARGET_AUDIENCE],
  };
};

export const bundleDataForOnlineApi = (item: IListItem): string => {
  const data = bundleItem(item);
  return JSON.stringify(data);
};

export const bundleDataForOnPrem = (
  item: IListItem,
  itemType: string
): string => {
  const data = bundleItem(item);
  const onPremData = { ...data, __metadata: { type: itemType } };
  return JSON.stringify(onPremData);
};

export const extractSpListItems = (spListItems: ISpListItem[]): IListItem[] => {
  const items: IListItem[] = spListItems.map((item) => {
    return {
      id: item[ID],
      courseName: item.Title,
      courseCode: item[COURSE_CODE],
      courseFrequency: item[COURSE_FREQUENCY],
      targetAudience: item[TARGET_AUDIENCE],
    };
  });
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
