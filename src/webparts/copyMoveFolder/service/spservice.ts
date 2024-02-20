// import { getSP } from "./pnpConfig";
// import "@pnp/sp/webs";
// import "@pnp/sp/files";
// import "@pnp/sp/folders";

// export interface IFolderWithFiles {
//   Name: string;
//   ServerRelativeUrl: string;
//   files: any[]; // Adjust the type of files according to your SharePoint setup
// }

// export const getAllFolders = async (): Promise<IFolderWithFiles[]> => {
//   try {
//     const sp = getSP();
//     const folders: any[] = await sp.web
//       .getFolderByServerRelativePath("RootFolder")
//       .folders();
//     const folderDetails: IFolderWithFiles[] = await Promise.all(
//       folders.map(async (folder) => {
//         const files: any[] = await folder.files();
//         return {
//           Name: folder.Name,
//           ServerRelativeUrl: folder.ServerRelativeUrl,
//           files,
//         };
//       })
//     );
//     return folderDetails;
//   } catch (error) {
//     console.error("Error retrieving folders:", error);
//     return [];
//   }
// };

// export const copyFolder = async (folder: IFolderWithFiles) => {
//   try {
//     const sp = getSP();
//     const copiedFolder = await sp.web
//       .getFolderByServerRelativePath("CopiedFolders")
//       .folders.addUsingPath(folder.Name);
//     const copiedFolderObject = await sp.web.getFolderByServerRelativePath(
//       copiedFolder.data.ServerRelativeUrl
//     ); // Fetch the folder object
//     await Promise.all(
//       folder.files.map(async (file) => {
//         await copiedFolderObject.files.addUsingPath(file.Name, file, {
//           Overwrite: true,
//         }); // Add file to the folder object
//         console.log(`File "${file.Name}" copied successfully.`);
//       })
//     );
//     console.log(`Folder "${folder.Name}" copied successfully.`);
//   } catch (error) {
//     console.error("Error copying folder:", error);
//     throw error;
//   }
// };

// export const moveFolder = async (folder: IFolderWithFiles) => {
//   try {
//     const sp = getSP();
//     const movedFolder = await sp.web
//       .getFolderByServerRelativePath("MovedFolders")
//       .folders.addUsingPath(folder.Name);
//     const movedFolderObject = await sp.web.getFolderByServerRelativePath(
//       movedFolder.data.ServerRelativeUrl
//     ); // Fetch the folder object
//     await Promise.all(
//       folder.files.map(async (file) => {
//         await movedFolderObject.files.addUsingPath(file.Name, file, {
//           Overwrite: true,
//         }); // Add file to the folder object
//         console.log(`File "${file.Name}" moved successfully.`);
//       })
//     );
//     await sp.web
//       .getFolderByServerRelativePath(folder.ServerRelativeUrl)
//       .recycle();
//     console.log(`Folder "${folder.Name}" moved successfully.`);
//     return true;
//   } catch (error) {
//     console.error("Error moving folder:", error);
//     throw error;
//   }
// };

import { getSP } from "./pnpConfig";
import "@pnp/sp/webs";
import "@pnp/sp/files";
import "@pnp/sp/folders";
import "@pnp/sp/presets/all";

export const getAllFolders = async () => {
  try {
    const sp = getSP();
    const rootFolder = await sp.web.getFolderByServerRelativePath("RootFolder");
    const allFolders = await rootFolder.folders();
    const sortedFolders = await allFolders.sort((a: any, b: any) =>
      a.Name.localeCompare(b.Name)
    );
    return sortedFolders;
  } catch (error) {
    console.error("Error retrieving folders:", error);
  }
};

export const copyFolder = async (folder: any) => {
  const sp = getSP();
  const destinationUrl = `/sites/QConnect/CopiedFolders/${folder.Name}`;
  await sp.web.rootFolder.folders
    .getByUrl("RootFolder")
    .folders.getByUrl(folder.Name)
    .copyByPath(destinationUrl, true);
};

export const moveFolder = async (folder: any) => {
  const sp = getSP();

  const destinationUrl = `/sites/QConnect/MovedFolders/${folder.Name}`;
  await sp.web.rootFolder.folders
    .getByUrl("RootFolder")
    .folders.getByUrl(folder.Name)
    .moveByPath(destinationUrl, true);
};
