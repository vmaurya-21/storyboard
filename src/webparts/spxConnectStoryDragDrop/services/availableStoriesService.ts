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
  /**
   * Display name of whoever added the story — SharePoint's built-in Created By.
   *
   * Used as the byline on carousel slides. Empty when the expand is
   * unavailable, which callers treat as "no byline" rather than an error.
   */
  author?: string;
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
  Description?: string;
  description?: string;
  /** Image hyperlink, as an object or a plain URL string. */
  ImageUrl?: { Url: string } | string;
  imageUrl?: { Url: string } | string;
  /** Post hyperlink, as an object or a plain URL string. */
  LinkToPost?: { Url: string } | string;
  linkToPost?: { Url: string } | string;
  /** Last-modified timestamp, surfaced as the display date. */
  Modified?: string;
  /** Creation timestamp. */
  Created?: string;
  /** Origin of the post. */
  Source?: string;
  /** Built-in Created By lookup, projected via `$expand=Author`. */
  Author?: { Title?: string };
}

/** Default title of the SharePoint list backing available stories. */
export const DEFAULT_AVAILABLE_STORIES_LIST_NAME = "AvailableStories";

let availableStoriesListName = DEFAULT_AVAILABLE_STORIES_LIST_NAME;
const WC_SHAREPOINT_HOST = "whitecasempsaemea.sharepoint.com";

/**
 * Override the SharePoint list title used by this service.
 *
 * Blank values are ignored and reset to the default.
 */
export const setAvailableStoriesListName = (listName?: string): void => {
  const next = (listName || "").trim();
  availableStoriesListName = next || DEFAULT_AVAILABLE_STORIES_LIST_NAME;
};

/** Current SharePoint list title used by this service. */
export const getAvailableStoriesListName = (): string => availableStoriesListName;

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

    if (hostname === WC_SHAREPOINT_HOST || hostname.endsWith(`.${WC_SHAREPOINT_HOST}`)) {
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
      Description: story.description || "",
    };

    payload.Source = deriveSourceFromUrl(story.linkToPost);

    if (story.imageUrl) {
      payload.ImageUrl = { Url: story.imageUrl };
    }
    if (story.linkToPost) {
      payload.LinkToPost = { Url: story.linkToPost };
    }

    const result = await client.web.lists.getByTitle(getAvailableStoriesListName()).items.add(payload);
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
      .getByTitle(getAvailableStoriesListName())
      .items.select(
        "ID",
        "Title",
        "Description",
        "ImageUrl",
        "LinkToPost",
        "Modified",
        "Created",
        "Source",
        // Created By, used as the story byline. Projecting a lookup's subfield
        // requires the matching $expand below.
        "Author/Title"
      )
      .expand("Author")();
    return items
      .map((item: ISharePointStoryItem) => {
      const itemImage = item.ImageUrl ?? item.imageUrl;
      const itemLink = item.LinkToPost ?? item.linkToPost;
      const imageUrl = typeof itemImage === 'object' ? itemImage?.Url || "" : (itemImage || "");
      const linkToPost = typeof itemLink === 'object' ? itemLink?.Url || "" : (itemLink || "");
      
      const rawSource = typeof item.Source === 'string' ? item.Source.trim().toLowerCase() : '';
      const sourceCategory = (rawSource === 'linkedin' || rawSource === 'external' || rawSource === 'internal')
        ? rawSource
        : deriveSourceFromUrl(linkToPost);

      return {
        id: item.ID,
        title: item.Title,
        description: item.Description ?? item.description ?? "",
        imageUrl: imageUrl,
        linkToPost: linkToPost,
        date: item.Modified || "",
        created: item.Created || "",
        source: sourceCategory,
        author: item.Author?.Title || "",
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
      Description: story.description || "",
    };

    payload.Source = deriveSourceFromUrl(story.linkToPost);

    if (story.imageUrl) {
      payload.ImageUrl = { Url: story.imageUrl };
    }
    if (story.linkToPost) {
      payload.LinkToPost = { Url: story.linkToPost };
    }

    await client.web.lists.getByTitle(getAvailableStoriesListName()).items.getById(id).update(payload);
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
    await client.web.lists.getByTitle(getAvailableStoriesListName()).items.getById(id).delete();
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
    const listName = getAvailableStoriesListName();
    const storyListExists = lists.some((list: ISharePointList) => list.Title === listName);

    if (storyListExists) {
      return;
    }

    await client.web.lists.add(listName);

    const list = client.web.lists.getByTitle(listName);
    
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
