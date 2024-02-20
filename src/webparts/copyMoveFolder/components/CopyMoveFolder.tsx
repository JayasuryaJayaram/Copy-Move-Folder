import * as React from "react";
import "@pnp/sp/folders";
import { ICopyMoveFolderProps } from "./ICopyMoveFolderProps";
import { copyFolder, getAllFolders, moveFolder } from "../service/spservice";
import { Modal } from "antd";
import { FolderOutlined } from "@ant-design/icons";
import styles from "./CopyMoveFolder.module.scss";

const CopyMoveFolder = (props: ICopyMoveFolderProps) => {
  const [folders, setFolders] = React.useState<any[]>([]);
  const [successModalVisible, setSuccessModalVisible] = React.useState(false);
  const [successMessage, setSuccessMessage] = React.useState("");

  React.useEffect(() => {
    const getFolders = async () => {
      try {
        const allFolders: any = await getAllFolders();
        // Sort the folders alphabetically by name

        setFolders(allFolders);
        console.log("Folders", allFolders);
      } catch (error) {
        console.error("Error retrieving folders:", error);
      }
    };

    getFolders();
  }, []);

  const handleCopy = async (folder: any) => {
    try {
      await copyFolder(folder);
      const updatedFolders: any = await getAllFolders();
      setFolders(updatedFolders);
      setSuccessMessage("Folder copied successfully");
      setSuccessModalVisible(true);
    } catch (error) {
      console.error("Error copying file:", error);
    }
  };

  const handleMove = async (folder: any) => {
    try {
      await moveFolder(folder);
      const updatedFolders: any = await getAllFolders();
      setFolders(updatedFolders);
      setSuccessMessage("Folder moved successfully");
      setSuccessModalVisible(true);
    } catch (error) {
      console.error("Error moving file:", error);
    }
  };

  const closeModal = () => {
    setSuccessModalVisible(false);
  };

  return (
    <div className={styles.container}>
      <div className={styles.header}>Folders in RootFolder</div>
      <div className={styles.foldersContainer}>
        {folders.map((folder) => (
          <div key={folder.UniqueId} className={styles.folderItem}>
            <div style={{ marginBottom: "10px" }}>
              {" "}
              <FolderOutlined rev={undefined} />
              {folder.Name}
            </div>
            <button
              onClick={() => handleCopy(folder)}
              className={styles.copyButton}
            >
              Copy
            </button>
            <button
              onClick={() => handleMove(folder)}
              className={styles.moveButton}
            >
              Move
            </button>
          </div>
        ))}
      </div>
      <Modal
        title="Success"
        visible={successModalVisible}
        onOk={closeModal}
        onCancel={closeModal}
      >
        {successMessage}
      </Modal>
    </div>
  );
};

export default CopyMoveFolder;
