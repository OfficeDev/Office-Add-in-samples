import { LeftOutlined, RightOutlined, DownOutlined } from "@ant-design/icons";
import { Button, Dropdown, MenuProps, Space } from "antd";
import React from "react";
import AIKeyConfigDialog from "./AIKeyConfigDialog";
import { apiKey, deployment, dropdownMenus, endpoint, generateText } from "./utility/AIData";
import {
  predefinedCitation,
  predefinedComment,
  predefinedDocumentTemplateBase64,
  predefinedPictureBase64,
  predefinedTitle,
} from "./utility/PredefinedData";

// global variable to store the EndPoint/Deployment/ApiKey, configrued by developer
export let _apiKey = "";
export let _endpoint = "";
export let _deployment = "";

export default class Home extends React.Component {
  constructor(props, context) {
    super(props, context);
  }

  state = {
    displayMainFunc: false,
    openKeyConfigDialog: false,
    titleLoading: false,
    commentLoading: false,
    citationLoading: false,
    pictureLoading: false,
    formatLoading: false,
    importTemplateLoading: false,
  };

  openMainFunc = () => {
    this.setState({ displayMainFunc: true });
  };

  back = () => {
    this.setState({ displayMainFunc: false });
  };

  open = (isOpen: boolean) => {
    this.setState({ openKeyConfigDialog: isOpen });
  };

  setKey = (key: string) => {
    _apiKey = key;
  };

  setEndpoint = (endpoint: string) => {
    _endpoint = endpoint;
  };

  setDeployment = (deployment: string) => {
    _deployment = deployment;
  };

  insertTemplateDocument = () => {
    return Word.run(async (context) => {
      this.setState({ importTemplateLoading: true });
      context.document.body.insertText("\n", Word.InsertLocation.end);
      const range = context.document.body.insertFileFromBase64(
        predefinedDocumentTemplateBase64,
        Word.InsertLocation.end
      );
      //locate the start position of the document
      range.getRange(Word.RangeLocation.start).select();
      await context.sync();
    })
      .catch(() => {
        //message.error(error.message);
      })
      .finally(() => {
        this.setState({ importTemplateLoading: false, displayMainFunc: true });
      });
  };

  //This is the code interacting with the Word document
  insertTitle = (titleStr: string) => {
    return Word.run(async (context) => {
      this.setState({ titleLoading: true });
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
      //locate the inserted title
      title.select();
      await context.sync();
    })
      .catch(() => {
        //message.error(error.message);
      })
      .finally(async () => {
        this.setState({ titleLoading: false });
      });
  };

  insertFootnote = (citation: string) => {
    return Word.run(async (context) => {
      this.setState({ citationLoading: true });
      const range = context.document.getSelection();
      const footnote = range.insertFootnote(citation);
      //locate the inserted footnote
      footnote.body.getRange().select();
      await context.sync();
    })
      .catch(() => {
        //message.error(error.message);
      })
      .finally(() => {
        this.setState({ citationLoading: false });
      });
  };

  insertComment = (commentStr: string) => {
    return Word.run(async (context) => {
      this.setState({ commentLoading: true });
      const range = context.document.getSelection();
      const comment = range.insertComment(commentStr);
      //locate the inserted comment
      comment.getRange().select();
      await context.sync();
    })
      .catch(() => {
        //message.error(error.message);
      })
      .finally(() => {
        this.setState({ commentLoading: false });
      });
  };

  insertPicture = (pictureBase64: string) => {
    return Word.run(async (context) => {
      this.setState({ pictureLoading: true });
      const range = context.document.getSelection();
      const picture = range.insertInlinePictureFromBase64(pictureBase64, Word.InsertLocation.start);
      //locate the inserted picture
      picture.getRange().select();
      await context.sync();
    })
      .catch(() => {
        //message.error(error.message);
      })
      .finally(() => {
        this.setState({ pictureLoading: false });
      });
  };

  formatDocument = () => {
    return Word.run(async (context) => {
      this.setState({ formatLoading: true });
      //set title to Heading 1 and text center alignment
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

      //unify the Headings to Heading2 and bold font
      const paragraphs = context.document.body.paragraphs;
      paragraphs.load();
      await context.sync();
      //skip the Title
      for (let i = 1; i < paragraphs.items.length; i++) {
        if (paragraphs.items[i].style == "Subtitle") {
          paragraphs.items[i].style = "Heading 2";
          paragraphs.items[i].font.bold = true;
        }
      }
      await context.sync();

      //set the list items of first level to bold
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

      //if there's pictures, set the pictures to be center alignment
      const pictures = context.document.body.inlinePictures;
      pictures.load();
      await context.sync();
      if (pictures.items.length > 0) {
        for (let k = 0; k < pictures.items.length; k++) {
          pictures.items[0].paragraph.alignment = "Centered";
          await context.sync();
        }
      }

      //if there's TBD or DONE keywords, set TBD to be red and DONE to be green
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
    })
      .catch(() => {
        //message.error(error.message);
      })
      .finally(() => {
        this.setState({ formatLoading: false });
      });
  };

  onMenuClick = async (e) => {
    if (
      (e.key === "titleAI" || e.key === "citationAI" || e.key === "commentAI") &&
      ((_apiKey === "" && apiKey === "") ||
        (_endpoint === "" && endpoint === "") ||
        (_deployment === "" && deployment === ""))
    ) {
      this.setState({ openKeyConfigDialog: true });
      return;
    }
    switch (e.key) {
      case "titleAI":
        await generateText("generate a title of meeting notes", 50).then((text) => {
          this.insertTitle(text);
        });
        break;
      case "titlePredefined":
        await this.insertTitle(predefinedTitle);
        break;
      case "citationAI":
        await generateText("generate a citation of meeting notes", 50).then((text) => {
          this.insertFootnote(text);
        });
        break;
      case "citationPredefined":
        await this.insertFootnote(predefinedCitation);
        break;
      case "commentAI":
        await generateText("generate a comment of meeting notes", 50).then((text) => {
          this.insertComment(text);
        });
        break;
      case "commentPredefined":
        await this.insertComment(predefinedComment);
        break;
    }
  };

  generateMenuItems = (type: string): MenuProps["items"] => {
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
          onClick: this.onMenuClick,
          selectable: true,
        };
      }
    });
  };

  render() {
    const addTitleItems: MenuProps["items"] = this.generateMenuItems("title");

    const addCitationItems: MenuProps["items"] = this.generateMenuItems("citation");

    const addCommentItems: MenuProps["items"] = this.generateMenuItems("comment");

    return (
      <>
        <div className="wrapper">
          <div className="main_content">
            {this.state.displayMainFunc ? (
              <>
                <div className="back">
                  <div className="cursor" onClick={this.back}>
                    <LeftOutlined />
                    <span>Back</span>
                  </div>
                </div>
                <div className="main_func">
                  <Dropdown menu={{ items: addTitleItems }} className="dropdown_list">
                    <Button loading={this.state.titleLoading}>
                      <Space>
                        Add a title
                        <DownOutlined />
                      </Space>
                    </Button>
                  </Dropdown>
                  <Dropdown menu={{ items: addCommentItems }} className="dropdown_list">
                    <Button loading={this.state.commentLoading}>
                      <Space>
                        Add a comment
                        <DownOutlined />
                      </Space>
                    </Button>
                  </Dropdown>
                  <Dropdown menu={{ items: addCitationItems }} className="dropdown_list">
                    <Button loading={this.state.citationLoading}>
                      <Space>
                        Add a footnote citation
                        <DownOutlined />
                      </Space>
                    </Button>
                  </Dropdown>
                  <Button
                    loading={this.state.pictureLoading}
                    className="insert_button"
                    onClick={() => this.insertPicture(predefinedPictureBase64)}
                  >
                    Add a sample image
                  </Button>
                  <Button loading={this.state.formatLoading} className="insert_button" onClick={this.formatDocument}>
                    Format the document
                  </Button>
                </div>
                <AIKeyConfigDialog
                  isOpen={this.state.openKeyConfigDialog}
                  endpoint={_endpoint}
                  deployment={_deployment}
                  apiKey={_apiKey}
                  setOpen={this.open.bind(this)}
                  setKey={this.setKey.bind(this)}
                  setEndpoint={this.setEndpoint.bind(this)}
                  setDeployment={this.setDeployment.bind(this)}
                />
              </>
            ) : (
              <>
                <div className="survey">
                  <RightOutlined />
                  <a
                    href="https://forms.office.com/Pages/ResponsePage.aspx?id=v4j5cvGGr0GRqy180BHbR8GFRbAYEV9Hmqgjcbr7lOdUNVAxQklNRkxCWEtMMFRFN0xXUFhYVlc5Ni4u"
                    target="_blank"
                    rel="noreferrer"
                  >
                    How do you like this sample? Tell us more!
                  </a>
                </div>
                <div className="header">
                  <div className="desc">
                    This sample add-in shows how to insert and format predefined or AI-generated content into a Word document.
                  </div>
                </div>
                <Button
                  className="generate_button"
                  onClick={this.insertTemplateDocument}
                  loading={this.state.importTemplateLoading}
                >
                  Generate sample content
                </Button>
                <div className="generate_button_or">or</div>
                <Button className="generate_button" onClick={this.openMainFunc}>
                  Create custom content
                </Button>
              </>
            )}
          </div>
          <div className="bottom">
            <div className="bottom_item">
              <RightOutlined className="item_icon" />
              <div className="bottom_item_info">
                <a
                  href="https://learn.microsoft.com/en-us/office/dev/add-ins/quickstarts/word-quickstart?tabs=yeomangenerator"
                  target="_blank"
                  rel="noreferrer"
                >
                  Explore more resources
                </a>
              </div>
            </div>
          </div>
        </div>
      </>
    );
  }
}
