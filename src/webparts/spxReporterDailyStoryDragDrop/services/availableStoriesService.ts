import { spfi, SPFx, SPFI } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/fields";
import { IWebPartContext } from "@microsoft/sp-webpart-base";
import { IItemAddResult } from "@pnp/sp/items";
let sp: SPFI | undefined;

/**
 * A story as this service exchanges it with callers.
 *
 * Field names are normalised to camel case, unlike the raw list item. Both
 * `id` and `ID` appear because SharePoint returns the latter while writes
 * accept the former.
 */
export interface IStoryItem {
  /** List item id, lower case as used on writes. */
  id?: number;
  /** List item id as SharePoint returns it on reads. */
  ID?: number;
  /** Headline. Maps to the list's `Title` column. */
  title: string;
  /** Optional summary body. */
  description?: string;
  /** Absolute image URL, flattened from the list's hyperlink column. */
  imageUrl: string;
  /** Absolute post URL, flattened from the list's hyperlink column. */
  linkToPost?: string;
  /** Display date, derived from the item's `Modified` timestamp. */
  date?: string;
  /** Raw `Created` timestamp, used for sorting. */
  created?: string;
  /** Origin of the post, from the list's `Source` column. */
  source?: string;
}

/**
 * Raw shape of an Available Stories list item.
 *
 * Hyperlink columns arrive either as `{ Url }` objects or as bare strings
 * depending on how the row was written, which is why both are modelled.
 */
interface ISharePointStoryItem {
  /** List item id. */
  ID: number;
  /** Headline. */
  Title: string;
  /** Optional summary body. */
  description?: string;
  /** Image hyperlink, as an object or a plain URL string. */
  imageUrl?: { Url: string } | string;
  /** Post hyperlink, as an object or a plain URL string. */
  linkToPost?: { Url: string } | string;
  /** Last-modified timestamp, surfaced as the display date. */
  Modified?: string;
  /** Creation timestamp. */
  Created?: string;
  /** Origin of the post. */
  Source?: string;
}

/** Title of the SharePoint list backing the available stories. */
const LIST_NAME = "Available Stories";

/**
 * Auto-detect a source category string ('internal' | 'external' | 'linkedin') from a post URL.
 */
export const deriveSourceFromUrl = (url?: string): 'internal' | 'external' | 'linkedin' => {
  if (!url || url === "#" || url.trim() === "") {
    return "internal";
  }
  try {
    const hostname = new URL(url).hostname.toLowerCase();

    if (hostname === "linkedin.com" || hostname.endsWith(".linkedin.com")) {
      return "linkedin";
    }

    if (hostname === "whitecase.com" || hostname.endsWith(".whitecase.com")) {
      if (hostname.startsWith("external.") || hostname.indexOf("external-") !== -1) {
        return "external";
      }
      return "internal";
    }

    return "external";
  } catch {
    return "internal";
  }
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
    };

    payload.Source = deriveSourceFromUrl(story.linkToPost);

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
        "Source"
      )();
    return items
      .map((item: ISharePointStoryItem) => {
      const imageUrl = typeof item.imageUrl === 'object' ? item.imageUrl?.Url || "" : (item.imageUrl || "");
      const linkToPost = typeof item.linkToPost === 'object' ? item.linkToPost?.Url || "" : (item.linkToPost || "");
      
      const rawSource = typeof item.Source === 'string' ? item.Source.trim().toLowerCase() : '';
      const sourceCategory = (rawSource === 'linkedin' || rawSource === 'external' || rawSource === 'internal')
        ? rawSource
        : deriveSourceFromUrl(linkToPost);

      return {
        id: item.ID,
        title: item.Title,
        description: item.description || "",
        imageUrl: imageUrl,
        linkToPost: linkToPost,
        date: item.Modified || "",
        created: item.Created || "",
        source: sourceCategory,
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

    payload.Source = deriveSourceFromUrl(story.linkToPost);

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
    await client.web.lists.getByTitle(LIST_NAME).items.getById(id).delete();
  } catch (error) {
    console.error("Error deleting story:", error);
    throw error;
  }
};

/**
 * Create the Stories list if it doesn't exist
 */
export const ensureListExists = async (): Promise<void> => {
  try {
    const client = getSpOrThrow();

    const lists = await client.web.lists();
    const storyListExists = lists.some((list: ISharePointList) => list.Title === LIST_NAME);

    if (storyListExists) {
      return;
    }

    await client.web.lists.add(LIST_NAME);

    const list = client.web.lists.getByTitle(LIST_NAME);
    
    await list.fields.addText("Description");

    await list.fields.addText("ImageUrl");

    await list.fields.addText("LinkToPost");

    await list.fields.addDateTime("StoryDate");

    await list.fields.addText("Source");
  } catch (error) {
    console.error("Error ensuring list exists:", error);
  }
};

interface ISharePointList {
  Title: string;
}
