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

interface IListItemPayload {
  id?: number;
  Title: string; // COURSE_NAME
  COURSE_CODE: string;
  COURSE_FREQUENCY: string;
  TARGET_AUDIENCE: string;
}

interface IListItemResponseOnPrem {
  data: {
    d: {
      results: any[];
    };
  };
}

const bundleItem = (item: IListItem): IListItemPayload => {
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

export const extractSpListItems = (
  spResponse: any,
  spOnline: boolean
): IListItem[] => {
  let items: IListItem[] = [];
  let listItems: any[] = [];

  if (spOnline) {
    if (spResponse?.value) {
      listItems = spResponse.value;
    }
  } else if (spResponse?.data?.d?.results?.length) {
    listItems = spResponse.data.d.results;
  }

  items = listItems.map((item) => {
    let itemData = undefined;
    let id = "Id";
    if (spOnline) {
      const onlineItem: IListItemPayload = item;
      itemData = onlineItem;
    } else {
      const onPremItem: IListItemResponseOnPrem = item;
      itemData = onPremItem;
    }
    if (itemData) {
      return {
        id: itemData[id],
        courseName: itemData.Title,
        courseCode: itemData[COURSE_CODE],
        courseFrequency: itemData[COURSE_FREQUENCY],
        targetAudience: itemData[TARGET_AUDIENCE],
      };
    }
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
