import { spfi, SPFx, SPFI } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/fields";
import { IWebPartContext } from "@microsoft/sp-webpart-base";
import { IItemAddResult } from "@pnp/sp/items";
let sp: SPFI | undefined;

export interface IStoryItem {
  id?: number;
  ID?: number;
  title: string;
  description?: string;
  imageUrl: string;
  linkToPost?: string;
  date?: string;
  created?: string;
}

interface ISharePointStoryItem {
  ID: number;
  Title: string;
  description?: string;
  imageUrl?: { Url: string } | string;
  linkToPost?: { Url: string } | string;
  Modified?: string; // SharePoint's built-in last-modified timestamp (ISO)
  Created?: string; // SharePoint's built-in creation timestamp (ISO)
  IsActive?: boolean | string; // Yes/No column — false/"No" means soft-deleted
}

const LIST_NAME = "Available Stories";

/**
 * Whether a story row should be visible. A story is active unless IsActive is
 * explicitly No/false. Missing values (e.g. rows created before the column was
 * added) are treated as active so they don't silently disappear.
 */
const isActiveStory = (value: boolean | string | undefined): boolean => {
  if (typeof value === "boolean") return value;
  if (typeof value === "string") {
    const v = value.trim().toLowerCase();
    return v === "yes" || v === "true" || v === "1";
  }
  return true;
};

/**
 * Get the current SPFI instance
 */
export const getSp = (): SPFI | undefined => {
  return sp;
};

/**
 * Return the initialized SPFI, or throw a clear error if it isn't ready.
 */
const getSpOrThrow = (): SPFI => {
  if (!sp || !sp.web) {
    throw new Error(
      "SharePoint not initialized. Please ensure initializeSharePoint is called with a valid context."
    );
  }
  return sp;
};

/**
 * Initialize SharePoint PnP
 */
export const initializeSharePoint = (spWeb: IWebPartContext): SPFI => {
  // Always reinitialize to ensure fresh context
  sp = spfi().using(SPFx(spWeb));
  return sp;
};

/**
 * Create a new story in SharePoint list
 */
export const createStory = async (story: IStoryItem): Promise<IItemAddResult> => {
  try {
    const client = getSpOrThrow();
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const payload: any = {
      Title: story.title,
      description: story.description || "",
      // New stories are active regardless of the column's default value.
      // IsActive is a Choice column, so it takes the string "Yes"/"No".
      IsActive: "Yes",
    };

    // Only include URL fields if they have
    if (story.imageUrl) {
      payload.imageUrl = { Url: story.imageUrl };
    }
    if (story.linkToPost) {
      payload.linkToPost = { Url: story.linkToPost };
    }

    const result = await client.web.lists.getByTitle(LIST_NAME).items.add(payload);
    return result;
  } catch (error) {
    console.error("Error creating story:", error);
    throw error;
  }
};

/**
 * Fetch all stories from SharePoint list
 */
export const getStories = async (): Promise<IStoryItem[]> => {
  try {
    const client = getSpOrThrow();
    const items = await client.web.lists
      .getByTitle(LIST_NAME)
      .items.select(
        "ID",
        "Title",
        "description",
        "imageUrl",
        "linkToPost",
        "Modified",
        "Created",
        "IsActive"
      )();
    return items
      .filter((item: ISharePointStoryItem) => isActiveStory(item.IsActive))
      .map((item: ISharePointStoryItem) => {
      // Extract URL from URL field objects if needed
      const imageUrl = typeof item.imageUrl === 'object' ? item.imageUrl?.Url || "" : (item.imageUrl || "");
      const linkToPost = typeof item.linkToPost === 'object' ? item.linkToPost?.Url || "" : (item.linkToPost || "");
      
      return {
        id: item.ID,
        title: item.Title,
        description: item.description || "",
        imageUrl: imageUrl,
        linkToPost: linkToPost,
        // Show the story's actual last-modified date, not the current date.
        date: item.Modified || "",
        created: item.Created || "",
      };
    });
  } catch (error) {
    console.error("Error fetching stories:", error);
    throw error;
  }
};

/**
 * Update a story in SharePoint list
 */
export const updateStory = async (id: number, story: IStoryItem): Promise<void> => {
  try {
    const client = getSpOrThrow();
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const payload: any = {
      Title: story.title,
      description: story.description || "",
    };
    
    // Only include URL fields if they have values
    if (story.imageUrl) {
      payload.imageUrl = { Url: story.imageUrl };
    }
    if (story.linkToPost) {
      payload.linkToPost = { Url: story.linkToPost };
    }
    
    await client.web.lists.getByTitle(LIST_NAME).items.getById(id).update(payload);
  } catch (error) {
    console.error("Error updating story:", error);
    throw error;
  }
};

/**
 * Delete a story from SharePoint list
 */
export const deleteStory = async (id: number): Promise<void> => {
  try {
    const client = getSpOrThrow();
    // Soft delete: mark the story inactive instead of removing it from the list.
    // getStories filters these out, so it disappears from the UI but is retained.
    // IsActive is a Choice column, so it takes the string "No".
    await client.web.lists.getByTitle(LIST_NAME).items.getById(id).update({ IsActive: "No" });
  } catch (error) {
    console.error("Error deactivating story:", error);
    throw error;
  }
};

/**
 * Create the Stories list if it doesn't exist
 */
export const ensureListExists = async (): Promise<void> => {
  try {
    const client = getSpOrThrow();

    // Check if list exists
    const lists = await client.web.lists();
    const storyListExists = lists.some((list: ISharePointList) => list.Title === LIST_NAME);

    if (storyListExists) {
      return; // List already exists
    }

    // Create new list
    await client.web.lists.add(LIST_NAME);

    // Add columns to the list
    const list = client.web.lists.getByTitle(LIST_NAME);
    
    // Add Description field
    await list.fields.addText("Description");

    // Add ImageUrl field
    await list.fields.addText("ImageUrl");

    // Add LinkToPost field
    await list.fields.addText("LinkToPost");

    // Add StoryDate field
    await list.fields.addDateTime("StoryDate");
  } catch (error) {
    console.error("Error ensuring list exists:", error);
  }
};

interface ISharePointList {
  Title: string;
}
