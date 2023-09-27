  export interface IListItem {
    id?: number;
    courseName: string;
    courseCode: string;
    courseFrequency: string;
    targetAudience: string;
}

export interface ISpListItem {
    id: string;
    webUrl: string;
    siteId: string;
    fields: any;
}

// export interface ISpListItemWritePayload {
//     fields: {
//         courseName: string;
//         courseCode: string;
//         courseFrequency: string;
//         targetAudience: string;
//     }
// }

export const SP_SITE = '8r1bcm.sharepoint.com';

// https://entra.microsoft.com/#view/Microsoft_AAD_IAM/TenantOverview.ReactView
export const TENANT_ID = 'bd04e98d-f006-4879-9451-b1096f9d1d03';

// https://8r1bcm.sharepoint.com/_layouts/15/listedit.aspx?List=%7B53eff35b-3e5a-4ef3-b87f-3baad80b982a%7D
export const LIST_ID = '53eff35b-3e5a-4ef3-b87f-3baad80b982a';

export const MS_GRAPH_URL_SP_SITE = 'https://graph.microsoft.com/v1.0/sites/' +
    SP_SITE +
    '/lists/' +
    LIST_ID +
    '/items?expand=fields&' // TODO $select=col-one,col-two,coln-n&

export const SP_LIST_URL = 'https://graph.microsoft.com/v1.0/sites/' +
    SP_SITE +    
    '/lists/' +
    LIST_ID +
    '/items';
  
    // '/_layouts/15/listedit.aspx?List=%7B53eff35b-3e5a-4ef3-b87f-3baad80b982a%7D';

//'https://8r1bcm.sharepoint.com/Lists/ccUsersTrainingCourses/items';
//'8r1bcm.sharepoint.com';

// export const LIST_COLS = [
//     'id',
//     'COURSE_NAME',
//     'COURSE_CODE',
//     'COURSE_FREQUENCY',
//     'TARGET_AUDIENCE',
// ];

// export interface IListItemPayload {
//     header?: any;
//     body: IListItem;
// }

// export function createItem(num: number) {

// }

export const COURSE_NAME = 'Course No ';
export const COURSE_CODE = 'Course Code ';

// export const newItem: IListItem = {
//     courseName: COURSE_NAME,
//     courseCode: COURSE_CODE,
//     courseFrequency: 'Annual',
//     targetAudience: 'PBO'
// }

const COL_MAP = {
    id: 'id',
    courseName: 'Title', // COURSE_NAME
    courseCode: 'COURSE_CODE',
    courseFrequency: 'COURSE_FREQUENCY',
    targetAudience: 'TARGET_AUDIENCE',
}

export function extractSpListItems(spListItems: ISpListItem[]): IListItem[] {
    const items: IListItem[] = [];
    if (spListItems && spListItems.length) {
        spListItems.forEach(spItem => {
            items.push({
                id: spItem.fields[COL_MAP.id],
                courseName: spItem.fields[COL_MAP.courseName],
                courseCode: spItem.fields[COL_MAP.courseCode],
                courseFrequency: spItem.fields[COL_MAP.courseFrequency],
                targetAudience: spItem.fields[COL_MAP.targetAudience],
            });
        });
    }
    return items;
}