/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

import React, { useState } from "react";
import { LeftOutlined, DownOutlined } from "@ant-design/icons";
import { Button, Dropdown, MenuProps, Space } from "antd";
import AIKeyConfigDialog from "./AIKeyConfigDialog";
import { dropdownMenus, generateText } from "./utility/AIData";
import {
  predefinedCitation,
  predefinedComment,
  predefinedDocumentTemplateBase64,
  predefinedPictureBase64,
  predefinedTitle,
} from "./utility/PredefinedData";

/* global Word, Office */

export default function Home() {
  const [displayMainFunc, setDisplayMainFunc] = useState(false);
  const [openKeyConfigDialog, setOpenKeyConfigDialog] = useState(false);
  const [titleLoading, setTitleLoading] = useState(false);
  const [commentLoading, setCommentLoading] = useState(false);
  const [citationLoading, setCitationLoading] = useState(false);
  const [pictureLoading, setPictureLoading] = useState(false);
  const [formatLoading, setFormatLoading] = useState(false);
  const [importTemplateLoading, setImportTemplateLoading] = useState(false);

  // Set the default values of the API key, endpoint, and deployment of Azure OpenAI service.
  const [apiKey, setApiKey] = useState("");
  const [endpoint, setEndpoint] = useState("");
  const [deployment, setDeployment] = useState("");

  const openMainFunc = () => {
    setDisplayMainFunc(true);
  };

  const back = () => {
    setDisplayMainFunc(false);
  };

  const open = (isOpen: boolean) => {
    setOpenKeyConfigDialog(isOpen);
  };

  const insertTemplateDocument = async () => {
    await Word.run(async (context) => {
      setImportTemplateLoading(true);
      context.document.body.insertText("\n", Word.InsertLocation.end);
      const range = context.document.body.insertFileFromBase64(
        predefinedDocumentTemplateBase64,
        Word.InsertLocation.end
      );
      // Locate the start position of the document.
      range.getRange(Word.RangeLocation.start).select();
      await context.sync();
    });
    setImportTemplateLoading(false);
    setDisplayMainFunc(true);
  };

  // This is the code interacting with the Word document.
  const insertTitle = async (titleStr: string) => {
    try {
      setTitleLoading(true);
      await Word.run(async (context) => {
        const title = context.document.body.insertParagraph(titleStr, Word.InsertLocation.start);
        let myLanguage = Office.context.displayLanguage;
        switch (myLanguage) {
          case "en-US":
            title.style = "Heading 1";
            break;
          case "fr-FR":
            title.style = "Titre 1";
            break;
          case "zh-CN":
            title.style = "标题 1";
            break;
        }
        title.select();
        await context.sync();
      });
    } finally {
      setTitleLoading(false);
    }
  };

  const insertFootnote = async (citation: string) => {
    try {
      setCitationLoading(true);
      await Word.run(async (context) => {
        const range = context.document.getSelection();
        const footnote = range.insertFootnote(citation);
        footnote.body.getRange().select();
        await context.sync();
      });
    } finally {
      setCitationLoading(false);
    }
  };

  const insertComment = async (commentStr: string) => {
    try {
      setCommentLoading(true);
      await Word.run(async (context) => {
        const range = context.document.getSelection();
        const comment = range.insertComment(commentStr);
        comment.getRange().select();
        await context.sync();
      });
    } finally {
      setCommentLoading(false);
    }
  };

  const insertPicture = async (pictureBase64: string) => {
    try {
      setPictureLoading(true);
      await Word.run(async (context) => {
        const range = context.document.getSelection();
        const picture = range.insertInlinePictureFromBase64(pictureBase64, Word.InsertLocation.start);
        picture.getRange().select();
        await context.sync();
      });
    } finally {
      setPictureLoading(false);
    }
  };

  const formatDocument = async () => {
    try {
      setFormatLoading(true);
      await Word.run(async (context) => {
        const firstPara = context.document.body.paragraphs.getFirst();
        let myLanguage = Office.context.displayLanguage;
        switch (myLanguage) {
          case "en-US":
            firstPara.style = "Heading 1";
            break;
          case "fr-FR":
            firstPara.style = "Titre 1";
            break;
          case "zh-CN":
            firstPara.style = "标题 1";
            break;
        }
        firstPara.alignment = "Centered";
        await context.sync();

        const paragraphs = context.document.body.paragraphs;
        paragraphs.load();
        await context.sync();
        for (let i = 1; i < paragraphs.items.length; i++) {
          if (paragraphs.items[i].style == "Subtitle") {
            paragraphs.items[i].style = "Heading 2";
            paragraphs.items[i].font.bold = true;
          }
        }
        await context.sync();

        const lists = context.document.body.lists;
        lists.load();
        await context.sync();
        for (let i = 0; i < lists.items.length; i++) {
          const list = lists.items[i];
          list.setLevelNumbering(0, Word.ListNumbering.upperRoman);
          const levelParas = list.getLevelParagraphs(0);
          levelParas.load();
          await context.sync();
          for (let j = 0; j < levelParas.items.length; j++) {
            const para = levelParas.items[j];
            para.font.bold = true;
          }
          await context.sync();
        }

        const pictures = context.document.body.inlinePictures;
        pictures.load();
        await context.sync();
        if (pictures.items.length > 0) {
          for (let k = 0; k < pictures.items.length; k++) {
            pictures.items[0].paragraph.alignment = "Centered";
            await context.sync();
          }
        }

        const tbdRanges = context.document.body.search("TBD", { matchCase: true });
        const doneRanges = context.document.body.search("DONE", { matchCase: true });
        tbdRanges.load();
        doneRanges.load();
        await context.sync();
        for (let i = 0; i < tbdRanges.items.length; i++) {
          tbdRanges.items[i].font.highlightColor = "yellow";
        }
        await context.sync();
        for (let i = 0; i < doneRanges.items.length; i++) {
          doneRanges.items[i].font.highlightColor = "Turquoise";
        }
        await context.sync();
      });
    } finally {
      setFormatLoading(false);
    }
  };

  const onMenuClick = async (e) => {
    if (
      (e.key === "titleAI" || e.key === "citationAI" || e.key === "commentAI") &&
      (apiKey === "" || endpoint === "" || deployment === "")
    ) {
      setOpenKeyConfigDialog(true);
      return;
    }
    switch (e.key) {
      case "titleAI": {
        const titleText = await generateText(apiKey, endpoint, deployment, "generate a title of meeting notes", 50);
        await insertTitle(titleText);
        break;
      }
      case "titlePredefined":
        await insertTitle(predefinedTitle);
        break;
      case "citationAI": {
        const citationText = await generateText(apiKey, endpoint, deployment, "generate a title of meeting notes", 50);
        await insertFootnote(citationText);
        break;
      }
      case "citationPredefined": {
        await insertFootnote(predefinedCitation);
        break;
      }
      case "commentAI": {
        const commentText = await generateText(apiKey, endpoint, deployment, "generate a title of meeting notes", 50);
        await insertComment(commentText);
        break;
      }
      case "commentPredefined": {
        await insertComment(predefinedComment);
        break;
      }
    }
  };

  const generateMenuItems = (type: string): MenuProps["items"] => {
    return dropdownMenus[type].map((item) => {
      if (item.type === "divider") {
        return { type: "divider" };
      } else {
        return {
          key: item.key,
          label: (
            <div style={{ textAlign: "center" }}>
              <span>{item.desc}</span>
            </div>
          ),
          onClick: onMenuClick,
          selectable: true,
        };
      }
    });
  };

  const addTitleItems: MenuProps["items"] = generateMenuItems("title");

  const addCitationItems: MenuProps["items"] = generateMenuItems("citation");

  const addCommentItems: MenuProps["items"] = generateMenuItems("comment");

  return (
    <>
      <div className="wrapper">
        <div className="main_content">
          {displayMainFunc ? (
            <>
              <div className="back">
                <div className="cursor" onClick={back}>
                  <LeftOutlined />
                  <span>Back</span>
                </div>
              </div>
              <div className="main_func">
                <Dropdown menu={{ items: addTitleItems }} className="dropdown_list">
                  <Button loading={titleLoading}>
                    <Space>
                      Add a title
                      <DownOutlined />
                    </Space>
                  </Button>
                </Dropdown>
                <Dropdown menu={{ items: addCommentItems }} className="dropdown_list">
                  <Button loading={commentLoading}>
                    <Space>
                      Add a comment
                      <DownOutlined />
                    </Space>
                  </Button>
                </Dropdown>
                <Dropdown menu={{ items: addCitationItems }} className="dropdown_list">
                  <Button loading={citationLoading}>
                    <Space>
                      Add a footnote citation
                      <DownOutlined />
                    </Space>
                  </Button>
                </Dropdown>
                <Button
                  loading={pictureLoading}
                  className="insert_button"
                  onClick={() => insertPicture(predefinedPictureBase64)}
                >
                  Add a sample image
                </Button>
                <Button loading={formatLoading} className="insert_button" onClick={formatDocument}>
                  Format the document
                </Button>
              </div>
              <AIKeyConfigDialog
                isOpen={openKeyConfigDialog}
                apiKey={apiKey}
                endpoint={endpoint}
                deployment={deployment}
                setOpen={open}
                setKey={setApiKey}
                setEndpoint={setEndpoint}
                setDeployment={setDeployment}
              />
            </>
          ) : (
            <>
              <div className="header">
                <div className="desc">
                  This sample add-in shows how to insert and format predefined or AI-generated content into a Word
                  document.
                </div>
              </div>
              <Button className="generate_button" onClick={insertTemplateDocument} loading={importTemplateLoading}>
                Generate sample content
              </Button>
              <div className="generate_button_or">or</div>
              <Button className="generate_button" onClick={openMainFunc}>
                Create custom content
              </Button>
            </>
          )}
        </div>
      </div>
    </>
  );
}
