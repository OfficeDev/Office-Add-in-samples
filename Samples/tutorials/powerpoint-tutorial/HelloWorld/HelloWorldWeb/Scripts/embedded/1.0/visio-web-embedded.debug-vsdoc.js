
var Visio;
(function (Visio) {
	var Application = (function(_super) {
		__extends(Application, _super);
		function Application() {
			/// <summary> Represents the Application. [Api set:  1.1] </summary>
			/// <field name="context" type="Visio.RequestContext">The request context associated with this object.</field>
			/// <field name="isNull" type="Boolean">Returns a boolean value for whether the corresponding object is null. You must call "context.sync()" before reading the isNull property.</field>
			/// <field name="showToolbars" type="Boolean">Show or Hide the standard toolbars. [Api set:  1.1]</field>
		}

		Application.prototype.load = function(option) {
			/// <summary>
			/// Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
			/// </summary>
			/// <param name="option" type="string | string[] | OfficeExtension.LoadOption"/>
			/// <returns type="Visio.Application"/>
		}

		Application.prototype.set = function() {
			/// <signature>
			/// <summary>Sets multiple properties on the object at the same time, based on JSON input.</summary>
			/// <param name="properties" type="Visio.Interfaces.ApplicationUpdateData">Properties described by the Visio.Interfaces.ApplicationUpdateData interface.</param>
			/// <param name="options" type="string">Options of the form { throwOnReadOnly?: boolean }
			/// <br />
			/// * throwOnReadOnly: Throw an error if the passed-in property list includes read-only properties (default = true).
			/// </param>
			/// </signature>
			/// <signature>
			/// <summary>Sets multiple properties on the object at the same time, based on an existing loaded object.</summary>
			/// <param name="properties" type="Application">An existing Application object, with properties that have already been loaded and synced.</param>
			/// </signature>
		}

		return Application;
	})(OfficeExtension.ClientObject);
	Visio.Application = Application;
})(Visio || (Visio = {__proto__: null}));

var Visio;
(function (Visio) {
	var Interfaces;
	(function (Interfaces) {
		var BoundingBox = (function() {
			function BoundingBox() {
				/// <summary> Represents the BoundingBox of the shape. [Api set:  1.1] </summary>
				/// <field name="height" type="Number">The distance between the top and bottom edges of the bounding box of the shape, excluding any data graphics associated with the shape. [Api set:  1.1]</field>
				/// <field name="width" type="Number">The distance between the left and right edges of the bounding box of the shape, excluding any data graphics associated with the shape. [Api set:  1.1]</field>
				/// <field name="x" type="Number">An integer that specifies the x-coordinate of the bounding box. [Api set:  1.1]</field>
				/// <field name="y" type="Number">An integer that specifies the y-coordinate of the bounding box. [Api set:  1.1]</field>
			}
			return BoundingBox;
		})();
		Interfaces.BoundingBox.__proto__ = null;
		Interfaces.BoundingBox = BoundingBox;
	})(Interfaces = Visio.Interfaces || (Visio.Interfaces = { __proto__: null}));
})(Visio || (Visio = {__proto__: null}));

var Visio;
(function (Visio) {
	var Comment = (function(_super) {
		__extends(Comment, _super);
		function Comment() {
			/// <summary> Represents the Comment. [Api set:  1.1] </summary>
			/// <field name="context" type="Visio.RequestContext">The request context associated with this object.</field>
			/// <field name="isNull" type="Boolean">Returns a boolean value for whether the corresponding object is null. You must call "context.sync()" before reading the isNull property.</field>
			/// <field name="author" type="String">A string that specifies the label of the shape data item. [Api set:  1.1]</field>
			/// <field name="date" type="String">A string that specifies the format of the shape data item. [Api set:  1.1]</field>
			/// <field name="text" type="String">A string that specifies the value of the shape data item. [Api set:  1.1]</field>
		}

		Comment.prototype.load = function(option) {
			/// <summary>
			/// Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
			/// </summary>
			/// <param name="option" type="string | string[] | OfficeExtension.LoadOption"/>
			/// <returns type="Visio.Comment"/>
		}

		Comment.prototype.set = function() {
			/// <signature>
			/// <summary>Sets multiple properties on the object at the same time, based on JSON input.</summary>
			/// <param name="properties" type="Visio.Interfaces.CommentUpdateData">Properties described by the Visio.Interfaces.CommentUpdateData interface.</param>
			/// <param name="options" type="string">Options of the form { throwOnReadOnly?: boolean }
			/// <br />
			/// * throwOnReadOnly: Throw an error if the passed-in property list includes read-only properties (default = true).
			/// </param>
			/// </signature>
			/// <signature>
			/// <summary>Sets multiple properties on the object at the same time, based on an existing loaded object.</summary>
			/// <param name="properties" type="Comment">An existing Comment object, with properties that have already been loaded and synced.</param>
			/// </signature>
		}

		return Comment;
	})(OfficeExtension.ClientObject);
	Visio.Comment = Comment;
})(Visio || (Visio = {__proto__: null}));

var Visio;
(function (Visio) {
	var CommentCollection = (function(_super) {
		__extends(CommentCollection, _super);
		function CommentCollection() {
			/// <summary> Represents the CommentCollection for a given Shape. [Api set:  1.1] </summary>
			/// <field name="context" type="Visio.RequestContext">The request context associated with this object.</field>
			/// <field name="isNull" type="Boolean">Returns a boolean value for whether the corresponding object is null. You must call "context.sync()" before reading the isNull property.</field>
			/// <field name="items" type="Array" elementType="Visio.Comment">Gets the loaded child items in this collection.</field>
		}

		CommentCollection.prototype.load = function(option) {
			/// <summary>
			/// Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
			/// </summary>
			/// <param name="option" type="string | string[] | OfficeExtension.LoadOption"/>
			/// <returns type="Visio.CommentCollection"/>
		}
		CommentCollection.prototype.getCount = function() {
			/// <summary>
			/// Gets the number of Shape Data Items. [Api set:  1.1]
			/// </summary>
			/// <returns type="OfficeExtension.ClientResult&lt;number&gt;"></returns>
			var result = new OfficeExtension.ClientResult();
			result.__proto__ = null;
			result.value = 0;
			return result;
		}
		CommentCollection.prototype.getItem = function(key) {
			/// <summary>
			/// Gets the Comment using its name. [Api set:  1.1]
			/// </summary>
			/// <param name="key" >Key is the name of the Comment to be retrieved.</param>
			/// <returns type="Visio.Comment"></returns>
		}

		return CommentCollection;
	})(OfficeExtension.ClientObject);
	Visio.CommentCollection = CommentCollection;
})(Visio || (Visio = {__proto__: null}));

var Visio;
(function (Visio) {
	var Interfaces;
	(function (Interfaces) {
		var DataRefreshCompleteEventArgs = (function() {
			function DataRefreshCompleteEventArgs() {
				/// <summary> Provides information about the document that raised the DataRefreshComplete event. [Api set:  1.1] </summary>
				/// <field name="document" type="Visio.Document">Gets the document object that raised the DataRefreshComplete event. [Api set:  1.1]</field>
				/// <field name="success" type="Boolean">Gets the success/failure of the DataRefreshComplete event. [Api set:  1.1]</field>
			}
			return DataRefreshCompleteEventArgs;
		})();
		Interfaces.DataRefreshCompleteEventArgs.__proto__ = null;
		Interfaces.DataRefreshCompleteEventArgs = DataRefreshCompleteEventArgs;
	})(Interfaces = Visio.Interfaces || (Visio.Interfaces = { __proto__: null}));
})(Visio || (Visio = {__proto__: null}));

var Visio;
(function (Visio) {
	var Document = (function(_super) {
		__extends(Document, _super);
		function Document() {
			/// <summary> Represents the Document class. [Api set:  1.1] </summary>
			/// <field name="context" type="Visio.RequestContext">The request context associated with this object.</field>
			/// <field name="isNull" type="Boolean">Returns a boolean value for whether the corresponding object is null. You must call "context.sync()" before reading the isNull property.</field>
			/// <field name="application" type="Visio.Application">Represents a Visio application instance that contains this document. Read-only. [Api set:  1.1]</field>
			/// <field name="pages" type="Visio.PageCollection">Represents a collection of pages associated with the document. Read-only. [Api set:  1.1]</field>
			/// <field name="view" type="Visio.DocumentView">Returns the DocumentView object. [Api set:  1.1]</field>
			/// <field name="onDataRefreshComplete" type="OfficeExtension.EventHandlers">Occurs when the data is refreshed in the diagram. [Api set:  1.1]</field>
			/// <field name="onPageLoadComplete" type="OfficeExtension.EventHandlers">Occurs when the page is finished loading. [Api set:  1.1]</field>
			/// <field name="onSelectionChanged" type="OfficeExtension.EventHandlers">Occurs when the current selection of shapes changes. [Api set:  1.1]</field>
			/// <field name="onShapeMouseEnter" type="OfficeExtension.EventHandlers">Occurs when the user moves the mouse pointer into the bounding box of a shape. [Api set:  1.1]</field>
			/// <field name="onShapeMouseLeave" type="OfficeExtension.EventHandlers">Occurs when the user moves the mouse out of the bounding box of a shape. [Api set:  1.1]</field>
		}

		Document.prototype.load = function(option) {
			/// <summary>
			/// Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
			/// </summary>
			/// <param name="option" type="string | string[] | OfficeExtension.LoadOption"/>
			/// <returns type="Visio.Document"/>
		}
		Document.prototype.getActivePage = function() {
			/// <summary>
			/// Returns the Active Page of the document. [Api set:  1.1]
			/// </summary>
			/// <returns type="Visio.Page"></returns>
		}
		Document.prototype.setActivePage = function(PageName) {
			/// <summary>
			/// Set the Active Page of the document. [Api set:  1.1]
			/// </summary>
			/// <param name="PageName" type="String">Name of the page</param>
			/// <returns ></returns>
		}
		Document.prototype.startDataRefresh = function() {
			/// <summary>
			/// Triggers the refresh of the data in the Diagram, for all pages. [Api set:  1.1]
			/// </summary>
			/// <returns ></returns>
		}
		Document.prototype.onDataRefreshComplete = {
			__proto__: null,
			add: function (handler) {
				/// <param name="handler" type="function(eventArgs: Visio.Interfaces.DataRefreshCompleteEventArgs)">Handler for the event. EventArgs: Provides information about the document that raised the DataRefreshComplete event. </param>
				/// <returns type="OfficeExtension.EventHandlerResult"></returns>
				var eventInfo = new Visio.Interfaces.DataRefreshCompleteEventArgs();
				eventInfo.__proto__ = null;
				handler(eventInfo);
			},
			remove: function (handler) {
				/// <param name="handler" type="function(eventArgs: Visio.Interfaces.DataRefreshCompleteEventArgs)">Handler for the event.</param>
				return;
			},
			removeAll: function () {
				return;
			}
		};
		Document.prototype.onPageLoadComplete = {
			__proto__: null,
			add: function (handler) {
				/// <param name="handler" type="function(eventArgs: Visio.Interfaces.PageLoadCompleteEventArgs)">Handler for the event. EventArgs: Provides information about the page that raised the PageLoadComplete event. </param>
				/// <returns type="OfficeExtension.EventHandlerResult"></returns>
				var eventInfo = new Visio.Interfaces.PageLoadCompleteEventArgs();
				eventInfo.__proto__ = null;
				handler(eventInfo);
			},
			remove: function (handler) {
				/// <param name="handler" type="function(eventArgs: Visio.Interfaces.PageLoadCompleteEventArgs)">Handler for the event.</param>
				return;
			},
			removeAll: function () {
				return;
			}
		};
		Document.prototype.onSelectionChanged = {
			__proto__: null,
			add: function (handler) {
				/// <param name="handler" type="function(eventArgs: Visio.Interfaces.SelectionChangedEventArgs)">Handler for the event. EventArgs: Provides information about the shape collection that raised the SelectionChanged event. </param>
				/// <returns type="OfficeExtension.EventHandlerResult"></returns>
				var eventInfo = new Visio.Interfaces.SelectionChangedEventArgs();
				eventInfo.__proto__ = null;
				handler(eventInfo);
			},
			remove: function (handler) {
				/// <param name="handler" type="function(eventArgs: Visio.Interfaces.SelectionChangedEventArgs)">Handler for the event.</param>
				return;
			},
			removeAll: function () {
				return;
			}
		};
		Document.prototype.onShapeMouseEnter = {
			__proto__: null,
			add: function (handler) {
				/// <param name="handler" type="function(eventArgs: Visio.Interfaces.ShapeMouseEnterEventArgs)">Handler for the event. EventArgs: Provides information about the shape that raised the ShapeMouseEnter event. </param>
				/// <returns type="OfficeExtension.EventHandlerResult"></returns>
				var eventInfo = new Visio.Interfaces.ShapeMouseEnterEventArgs();
				eventInfo.__proto__ = null;
				handler(eventInfo);
			},
			remove: function (handler) {
				/// <param name="handler" type="function(eventArgs: Visio.Interfaces.ShapeMouseEnterEventArgs)">Handler for the event.</param>
				return;
			},
			removeAll: function () {
				return;
			}
		};
		Document.prototype.onShapeMouseLeave = {
			__proto__: null,
			add: function (handler) {
				/// <param name="handler" type="function(eventArgs: Visio.Interfaces.ShapeMouseLeaveEventArgs)">Handler for the event. EventArgs: Provides information about the shape that raised the ShapeMouseLeave event. </param>
				/// <returns type="OfficeExtension.EventHandlerResult"></returns>
				var eventInfo = new Visio.Interfaces.ShapeMouseLeaveEventArgs();
				eventInfo.__proto__ = null;
				handler(eventInfo);
			},
			remove: function (handler) {
				/// <param name="handler" type="function(eventArgs: Visio.Interfaces.ShapeMouseLeaveEventArgs)">Handler for the event.</param>
				return;
			},
			removeAll: function () {
				return;
			}
		};

		return Document;
	})(OfficeExtension.ClientObject);
	Visio.Document = Document;
})(Visio || (Visio = {__proto__: null}));

var Visio;
(function (Visio) {
	var DocumentView = (function(_super) {
		__extends(DocumentView, _super);
		function DocumentView() {
			/// <summary> Represents the DocumentView class. [Api set:  1.1] </summary>
			/// <field name="context" type="Visio.RequestContext">The request context associated with this object.</field>
			/// <field name="isNull" type="Boolean">Returns a boolean value for whether the corresponding object is null. You must call "context.sync()" before reading the isNull property.</field>
			/// <field name="disableHyperlinks" type="Boolean">Disable Hyperlinks. [Api set:  1.1]</field>
			/// <field name="disablePan" type="Boolean">Disable Pan. [Api set:  1.1]</field>
			/// <field name="disableZoom" type="Boolean">Disable Zoom. [Api set:  1.1]</field>
			/// <field name="hideDiagramBoundary" type="Boolean">Disable Hyperlinks. [Api set:  1.1]</field>
		}

		DocumentView.prototype.load = function(option) {
			/// <summary>
			/// Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
			/// </summary>
			/// <param name="option" type="string | string[] | OfficeExtension.LoadOption"/>
			/// <returns type="Visio.DocumentView"/>
		}

		DocumentView.prototype.set = function() {
			/// <signature>
			/// <summary>Sets multiple properties on the object at the same time, based on JSON input.</summary>
			/// <param name="properties" type="Visio.Interfaces.DocumentViewUpdateData">Properties described by the Visio.Interfaces.DocumentViewUpdateData interface.</param>
			/// <param name="options" type="string">Options of the form { throwOnReadOnly?: boolean }
			/// <br />
			/// * throwOnReadOnly: Throw an error if the passed-in property list includes read-only properties (default = true).
			/// </param>
			/// </signature>
			/// <signature>
			/// <summary>Sets multiple properties on the object at the same time, based on an existing loaded object.</summary>
			/// <param name="properties" type="DocumentView">An existing DocumentView object, with properties that have already been loaded and synced.</param>
			/// </signature>
		}

		return DocumentView;
	})(OfficeExtension.ClientObject);
	Visio.DocumentView = DocumentView;
})(Visio || (Visio = {__proto__: null}));

var Visio;
(function (Visio) {
	var Interfaces;
	(function (Interfaces) {
		var Highlight = (function() {
			function Highlight() {
				/// <summary> Represents the highlight data added to the shape. [Api set:  1.1] </summary>
				/// <field name="color" type="String">A string that specifies the color of the highlight. It must have the form &quot;#RRGGBB&quot;, where each letter represents a hexadecimal digit between 0 and F, and where RR is the red value between 0 and 0xFF (255), GG the green value between 0 and 0xFF (255), and BB is the blue value between 0 and 0xFF (255). [Api set:  1.1]</field>
				/// <field name="width" type="Number">A positive integer that specifies the width of the highlight&apos;s stroke in pixels. [Api set:  1.1]</field>
			}
			return Highlight;
		})();
		Interfaces.Highlight.__proto__ = null;
		Interfaces.Highlight = Highlight;
	})(Interfaces = Visio.Interfaces || (Visio.Interfaces = { __proto__: null}));
})(Visio || (Visio = {__proto__: null}));

var Visio;
(function (Visio) {
	var Hyperlink = (function(_super) {
		__extends(Hyperlink, _super);
		function Hyperlink() {
			/// <summary> Represents the Hyperlink. [Api set:  1.1] </summary>
			/// <field name="context" type="Visio.RequestContext">The request context associated with this object.</field>
			/// <field name="isNull" type="Boolean">Returns a boolean value for whether the corresponding object is null. You must call "context.sync()" before reading the isNull property.</field>
			/// <field name="address" type="String">Gets the address of the Hyperlink object. [Api set:  1.1]</field>
			/// <field name="description" type="String">Gets the description of a hyperlink. [Api set:  1.1]</field>
			/// <field name="subAddress" type="String">Gets the sub-address of the Hyperlink object. [Api set:  1.1]</field>
		}

		Hyperlink.prototype.load = function(option) {
			/// <summary>
			/// Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
			/// </summary>
			/// <param name="option" type="string | string[] | OfficeExtension.LoadOption"/>
			/// <returns type="Visio.Hyperlink"/>
		}

		return Hyperlink;
	})(OfficeExtension.ClientObject);
	Visio.Hyperlink = Hyperlink;
})(Visio || (Visio = {__proto__: null}));

var Visio;
(function (Visio) {
	var HyperlinkCollection = (function(_super) {
		__extends(HyperlinkCollection, _super);
		function HyperlinkCollection() {
			/// <summary> Represents the Hyperlink Collection. [Api set:  1.1] </summary>
			/// <field name="context" type="Visio.RequestContext">The request context associated with this object.</field>
			/// <field name="isNull" type="Boolean">Returns a boolean value for whether the corresponding object is null. You must call "context.sync()" before reading the isNull property.</field>
			/// <field name="items" type="Array" elementType="Visio.Hyperlink">Gets the loaded child items in this collection.</field>
		}

		HyperlinkCollection.prototype.load = function(option) {
			/// <summary>
			/// Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
			/// </summary>
			/// <param name="option" type="string | string[] | OfficeExtension.LoadOption"/>
			/// <returns type="Visio.HyperlinkCollection"/>
		}
		HyperlinkCollection.prototype.getCount = function() {
			/// <summary>
			/// Gets the number of hyperlinks. [Api set:  1.1]
			/// </summary>
			/// <returns type="OfficeExtension.ClientResult&lt;number&gt;"></returns>
			var result = new OfficeExtension.ClientResult();
			result.__proto__ = null;
			result.value = 0;
			return result;
		}
		HyperlinkCollection.prototype.getItem = function(Key) {
			/// <summary>
			/// Gets a Hyperlink using its key (name or Id). [Api set:  1.1]
			/// </summary>
			/// <param name="Key" >Key is the name or index of the Hyperlink to be retrieved.</param>
			/// <returns type="Visio.Hyperlink"></returns>
		}

		return HyperlinkCollection;
	})(OfficeExtension.ClientObject);
	Visio.HyperlinkCollection = HyperlinkCollection;
})(Visio || (Visio = {__proto__: null}));

var Visio;
(function (Visio) {
	/// <summary> Represents the Horizontal Alignment of the Overlay relative to the shape. [Api set:  1.1] </summary>
	var OverlayHorizontalAlignment = {
		__proto__: null,
		"left": "left",
		"center": "center",
		"right": "right",
	}
	Visio.OverlayHorizontalAlignment = OverlayHorizontalAlignment;
})(Visio || (Visio = {__proto__: null}));

var Visio;
(function (Visio) {
	/// <summary> Represents the type of the overlay. [Api set:  1.1] </summary>
	var OverlayType = {
		__proto__: null,
		"text": "text",
		"image": "image",
	}
	Visio.OverlayType = OverlayType;
})(Visio || (Visio = {__proto__: null}));

var Visio;
(function (Visio) {
	/// <summary> Represents the Vertical Alignment of the Overlay relative to the shape. [Api set:  1.1] </summary>
	var OverlayVerticalAlignment = {
		__proto__: null,
		"top": "top",
		"middle": "middle",
		"bottom": "bottom",
	}
	Visio.OverlayVerticalAlignment = OverlayVerticalAlignment;
})(Visio || (Visio = {__proto__: null}));

var Visio;
(function (Visio) {
	var Page = (function(_super) {
		__extends(Page, _super);
		function Page() {
			/// <summary> Represents the Page class. [Api set:  1.1] </summary>
			/// <field name="context" type="Visio.RequestContext">The request context associated with this object.</field>
			/// <field name="isNull" type="Boolean">Returns a boolean value for whether the corresponding object is null. You must call "context.sync()" before reading the isNull property.</field>
			/// <field name="allShapes" type="Visio.ShapeCollection">All shapes in the page. Read-only. [Api set:  1.1]</field>
			/// <field name="comments" type="Visio.CommentCollection">Returns the Comments Collection [Api set:  1.1]</field>
			/// <field name="height" type="Number">Returns the height of the page. Read-only. [Api set:  1.1]</field>
			/// <field name="index" type="Number">Index of the Page. [Api set:  1.1]</field>
			/// <field name="isBackground" type="Boolean">Whether the page is a background page or not. Read-only. [Api set:  1.1]</field>
			/// <field name="name" type="String">Page name. Read-only. [Api set:  1.1]</field>
			/// <field name="shapes" type="Visio.ShapeCollection">Shapes at root level, in the page. Read-only. [Api set:  1.1]</field>
			/// <field name="view" type="Visio.PageView">Returns the view of the page. Read-only. [Api set:  1.1]</field>
			/// <field name="width" type="Number">Returns the width of the page. Read-only. [Api set:  1.1]</field>
		}

		Page.prototype.load = function(option) {
			/// <summary>
			/// Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
			/// </summary>
			/// <param name="option" type="string | string[] | OfficeExtension.LoadOption"/>
			/// <returns type="Visio.Page"/>
		}
		Page.prototype.activate = function() {
			/// <summary>
			/// Set the page as Active Page of the document. [Api set:  1.1]
			/// </summary>
			/// <returns ></returns>
		}

		return Page;
	})(OfficeExtension.ClientObject);
	Visio.Page = Page;
})(Visio || (Visio = {__proto__: null}));

var Visio;
(function (Visio) {
	var PageCollection = (function(_super) {
		__extends(PageCollection, _super);
		function PageCollection() {
			/// <summary> Represents a collection of Page objects that are part of the document. [Api set:  1.1] </summary>
			/// <field name="context" type="Visio.RequestContext">The request context associated with this object.</field>
			/// <field name="isNull" type="Boolean">Returns a boolean value for whether the corresponding object is null. You must call "context.sync()" before reading the isNull property.</field>
			/// <field name="items" type="Array" elementType="Visio.Page">Gets the loaded child items in this collection.</field>
		}

		PageCollection.prototype.load = function(option) {
			/// <summary>
			/// Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
			/// </summary>
			/// <param name="option" type="string | string[] | OfficeExtension.LoadOption"/>
			/// <returns type="Visio.PageCollection"/>
		}
		PageCollection.prototype.getCount = function() {
			/// <summary>
			/// Gets the number of pages in the collection. [Api set:  1.1]
			/// </summary>
			/// <returns type="OfficeExtension.ClientResult&lt;number&gt;"></returns>
			var result = new OfficeExtension.ClientResult();
			result.__proto__ = null;
			result.value = 0;
			return result;
		}
		PageCollection.prototype.getItem = function(key) {
			/// <summary>
			/// Gets a page using its key (name or Id). [Api set:  1.1]
			/// </summary>
			/// <param name="key" >Key is the name or Id of the page to be retrieved.</param>
			/// <returns type="Visio.Page"></returns>
		}

		return PageCollection;
	})(OfficeExtension.ClientObject);
	Visio.PageCollection = PageCollection;
})(Visio || (Visio = {__proto__: null}));

var Visio;
(function (Visio) {
	var Interfaces;
	(function (Interfaces) {
		var PageLoadCompleteEventArgs = (function() {
			function PageLoadCompleteEventArgs() {
				/// <summary> Provides information about the page that raised the PageLoadComplete event. [Api set:  1.1] </summary>
				/// <field name="pageName" type="String">Gets the name of the page that raised the PageLoad event. [Api set:  1.1]</field>
				/// <field name="success" type="Boolean">Gets the success/failure of the PageLoadComplete event. [Api set:  1.1]</field>
			}
			return PageLoadCompleteEventArgs;
		})();
		Interfaces.PageLoadCompleteEventArgs.__proto__ = null;
		Interfaces.PageLoadCompleteEventArgs = PageLoadCompleteEventArgs;
	})(Interfaces = Visio.Interfaces || (Visio.Interfaces = { __proto__: null}));
})(Visio || (Visio = {__proto__: null}));

var Visio;
(function (Visio) {
	var PageView = (function(_super) {
		__extends(PageView, _super);
		function PageView() {
			/// <summary> Represents the PageView class. [Api set:  1.1] </summary>
			/// <field name="context" type="Visio.RequestContext">The request context associated with this object.</field>
			/// <field name="isNull" type="Boolean">Returns a boolean value for whether the corresponding object is null. You must call "context.sync()" before reading the isNull property.</field>
			/// <field name="zoom" type="Number">Get/Set Page&apos;s Zoom level. The value can be between 10 and 400 and denotes the percentage of zoom. [Api set:  1.1]</field>
		}

		PageView.prototype.load = function(option) {
			/// <summary>
			/// Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
			/// </summary>
			/// <param name="option" type="string | string[] | OfficeExtension.LoadOption"/>
			/// <returns type="Visio.PageView"/>
		}

		PageView.prototype.set = function() {
			/// <signature>
			/// <summary>Sets multiple properties on the object at the same time, based on JSON input.</summary>
			/// <param name="properties" type="Visio.Interfaces.PageViewUpdateData">Properties described by the Visio.Interfaces.PageViewUpdateData interface.</param>
			/// <param name="options" type="string">Options of the form { throwOnReadOnly?: boolean }
			/// <br />
			/// * throwOnReadOnly: Throw an error if the passed-in property list includes read-only properties (default = true).
			/// </param>
			/// </signature>
			/// <signature>
			/// <summary>Sets multiple properties on the object at the same time, based on an existing loaded object.</summary>
			/// <param name="properties" type="PageView">An existing PageView object, with properties that have already been loaded and synced.</param>
			/// </signature>
		}
		PageView.prototype.centerViewportOnShape = function(ShapeId) {
			/// <summary>
			/// Pans the Visio drawing to place the specified shape in the center of the view. [Api set:  1.1]
			/// </summary>
			/// <param name="ShapeId" type="Number">ShapeId to be seen in the center.</param>
			/// <returns ></returns>
		}
		PageView.prototype.fitToWindow = function() {
			/// <summary>
			/// Fit Page to current window. [Api set:  1.1]
			/// </summary>
			/// <returns ></returns>
		}
		PageView.prototype.getPosition = function() {
			/// <summary>
			/// Returns the position object that specifies the position of the page in the view. [Api set:  1.1]
			/// </summary>
			/// <returns type="OfficeExtension.ClientResult&lt;Visio.Position&gt;"></returns>
			var result = new OfficeExtension.ClientResult();
			result.__proto__ = null;
			result.value = {};
			return result;
		}
		PageView.prototype.getSelection = function() {
			/// <summary>
			/// Represents the Selection in the page. [Api set:  1.1]
			/// </summary>
			/// <returns type="Visio.Selection"></returns>
		}
		PageView.prototype.isShapeInViewport = function(Shape) {
			/// <summary>
			/// To check if the shape is in view of the page or not. [Api set:  1.1]
			/// </summary>
			/// <param name="Shape" type="Visio.Shape">Shape to be checked.</param>
			/// <returns type="OfficeExtension.ClientResult&lt;boolean&gt;"></returns>
			var result = new OfficeExtension.ClientResult();
			result.__proto__ = null;
			result.value = false;
			return result;
		}
		PageView.prototype.setPosition = function(Position) {
			/// <summary>
			/// Sets the position of the page in the view. [Api set:  1.1]
			/// </summary>
			/// <param name="Position" type="Visio.Interfaces.Position">Position object that specifies the new position of the page in the view.</param>
			/// <returns ></returns>
		}

		return PageView;
	})(OfficeExtension.ClientObject);
	Visio.PageView = PageView;
})(Visio || (Visio = {__proto__: null}));

var Visio;
(function (Visio) {
	var Interfaces;
	(function (Interfaces) {
		var Position = (function() {
			function Position() {
				/// <summary> Represents the Position of the object in the view. [Api set:  1.1] </summary>
				/// <field name="x" type="Number">An integer that specifies the x-coordinate of the object, which is the signed value of the distance in pixels from the viewport&apos;s center to the left boundary of the page. [Api set:  1.1]</field>
				/// <field name="y" type="Number">An integer that specifies the y-coordinate of the object, which is the signed value of the distance in pixels from the viewport&apos;s center to the top boundary of the page. [Api set:  1.1]</field>
			}
			return Position;
		})();
		Interfaces.Position.__proto__ = null;
		Interfaces.Position = Position;
	})(Interfaces = Visio.Interfaces || (Visio.Interfaces = { __proto__: null}));
})(Visio || (Visio = {__proto__: null}));

var Visio;
(function (Visio) {
	var Selection = (function(_super) {
		__extends(Selection, _super);
		function Selection() {
			/// <summary> Represents the Selection in the page. [Api set:  1.1] </summary>
			/// <field name="context" type="Visio.RequestContext">The request context associated with this object.</field>
			/// <field name="isNull" type="Boolean">Returns a boolean value for whether the corresponding object is null. You must call "context.sync()" before reading the isNull property.</field>
			/// <field name="shapes" type="Visio.ShapeCollection">Gets the Shapes of the Selection [Api set:  1.1]</field>
		}

		Selection.prototype.load = function(option) {
			/// <summary>
			/// Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
			/// </summary>
			/// <param name="option" type="string | string[] | OfficeExtension.LoadOption"/>
			/// <returns type="Visio.Selection"/>
		}

		return Selection;
	})(OfficeExtension.ClientObject);
	Visio.Selection = Selection;
})(Visio || (Visio = {__proto__: null}));

var Visio;
(function (Visio) {
	var Interfaces;
	(function (Interfaces) {
		var SelectionChangedEventArgs = (function() {
			function SelectionChangedEventArgs() {
				/// <summary> Provides information about the shape collection that raised the SelectionChanged event. [Api set:  1.1] </summary>
				/// <field name="pageName" type="String">Gets the name of the page which has the ShapeCollection object that raised the SelectionChanged event. [Api set:  1.1]</field>
				/// <field name="shapeNames" type="Array" elementType="String">Gets the ShapeCollection object that raised the SelectionChanged event. [Api set:  1.1]</field>
			}
			return SelectionChangedEventArgs;
		})();
		Interfaces.SelectionChangedEventArgs.__proto__ = null;
		Interfaces.SelectionChangedEventArgs = SelectionChangedEventArgs;
	})(Interfaces = Visio.Interfaces || (Visio.Interfaces = { __proto__: null}));
})(Visio || (Visio = {__proto__: null}));

var Visio;
(function (Visio) {
	var Shape = (function(_super) {
		__extends(Shape, _super);
		function Shape() {
			/// <summary> Represents the Shape class. [Api set:  1.1] </summary>
			/// <field name="context" type="Visio.RequestContext">The request context associated with this object.</field>
			/// <field name="isNull" type="Boolean">Returns a boolean value for whether the corresponding object is null. You must call "context.sync()" before reading the isNull property.</field>
			/// <field name="comments" type="Visio.CommentCollection">Returns the Comments Collection [Api set:  1.1]</field>
			/// <field name="hyperlinks" type="Visio.HyperlinkCollection">Returns the Hyperlinks collection for a Shape object. Read-only. [Api set:  1.1]</field>
			/// <field name="id" type="Number">Shape&apos;s Identifier. [Api set:  1.1]</field>
			/// <field name="name" type="String">Shape&apos;s name. [Api set:  1.1]</field>
			/// <field name="select" type="Boolean">Returns true, if shape is selected. User can set true to select the shape explicitly. [Api set:  1.1]</field>
			/// <field name="shapeDataItems" type="Visio.ShapeDataItemCollection">Returns the Shape&apos;s Data Section. Read-only. [Api set:  1.1]</field>
			/// <field name="subShapes" type="Visio.ShapeCollection">Gets SubShape Collection. [Api set:  1.1]</field>
			/// <field name="text" type="String">Shape&apos;s Text. [Api set:  1.1]</field>
			/// <field name="view" type="Visio.ShapeView">Returns the view of the shape. Read-only. [Api set:  1.1]</field>
		}

		Shape.prototype.load = function(option) {
			/// <summary>
			/// Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
			/// </summary>
			/// <param name="option" type="string | string[] | OfficeExtension.LoadOption"/>
			/// <returns type="Visio.Shape"/>
		}

		Shape.prototype.set = function() {
			/// <signature>
			/// <summary>Sets multiple properties on the object at the same time, based on JSON input.</summary>
			/// <param name="properties" type="Visio.Interfaces.ShapeUpdateData">Properties described by the Visio.Interfaces.ShapeUpdateData interface.</param>
			/// <param name="options" type="string">Options of the form { throwOnReadOnly?: boolean }
			/// <br />
			/// * throwOnReadOnly: Throw an error if the passed-in property list includes read-only properties (default = true).
			/// </param>
			/// </signature>
			/// <signature>
			/// <summary>Sets multiple properties on the object at the same time, based on an existing loaded object.</summary>
			/// <param name="properties" type="Shape">An existing Shape object, with properties that have already been loaded and synced.</param>
			/// </signature>
		}
		Shape.prototype.getBounds = function() {
			/// <summary>
			/// Returns the BoundingBox object that specifies bounding box of the shape. [Api set:  1.1]
			/// </summary>
			/// <returns type="OfficeExtension.ClientResult&lt;Visio.BoundingBox&gt;"></returns>
			var result = new OfficeExtension.ClientResult();
			result.__proto__ = null;
			result.value = {};
			return result;
		}

		return Shape;
	})(OfficeExtension.ClientObject);
	Visio.Shape = Shape;
})(Visio || (Visio = {__proto__: null}));

var Visio;
(function (Visio) {
	var ShapeCollection = (function(_super) {
		__extends(ShapeCollection, _super);
		function ShapeCollection() {
			/// <summary> Represents the Shape Collection. [Api set:  1.1] </summary>
			/// <field name="context" type="Visio.RequestContext">The request context associated with this object.</field>
			/// <field name="isNull" type="Boolean">Returns a boolean value for whether the corresponding object is null. You must call "context.sync()" before reading the isNull property.</field>
			/// <field name="items" type="Array" elementType="Visio.Shape">Gets the loaded child items in this collection.</field>
		}

		ShapeCollection.prototype.load = function(option) {
			/// <summary>
			/// Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
			/// </summary>
			/// <param name="option" type="string | string[] | OfficeExtension.LoadOption"/>
			/// <returns type="Visio.ShapeCollection"/>
		}
		ShapeCollection.prototype.getCount = function() {
			/// <summary>
			/// Gets the number of Shapes in the collection. [Api set:  1.1]
			/// </summary>
			/// <returns type="OfficeExtension.ClientResult&lt;number&gt;"></returns>
			var result = new OfficeExtension.ClientResult();
			result.__proto__ = null;
			result.value = 0;
			return result;
		}
		ShapeCollection.prototype.getItem = function(key) {
			/// <summary>
			/// Gets a Shape using its key (name or Index). [Api set:  1.1]
			/// </summary>
			/// <param name="key" >Key is the Name or Index of the shape to be retrieved.</param>
			/// <returns type="Visio.Shape"></returns>
		}

		return ShapeCollection;
	})(OfficeExtension.ClientObject);
	Visio.ShapeCollection = ShapeCollection;
})(Visio || (Visio = {__proto__: null}));

var Visio;
(function (Visio) {
	var ShapeDataItem = (function(_super) {
		__extends(ShapeDataItem, _super);
		function ShapeDataItem() {
			/// <summary> Represents the ShapeDataItem. [Api set:  1.1] </summary>
			/// <field name="context" type="Visio.RequestContext">The request context associated with this object.</field>
			/// <field name="isNull" type="Boolean">Returns a boolean value for whether the corresponding object is null. You must call "context.sync()" before reading the isNull property.</field>
			/// <field name="format" type="String">A string that specifies the format of the shape data item. [Api set:  1.1]</field>
			/// <field name="formattedValue" type="String">A string that specifies the formatted value of the shape data item. [Api set:  1.1]</field>
			/// <field name="label" type="String">A string that specifies the label of the shape data item. [Api set:  1.1]</field>
			/// <field name="value" type="String">A string that specifies the value of the shape data item. [Api set:  1.1]</field>
		}

		ShapeDataItem.prototype.load = function(option) {
			/// <summary>
			/// Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
			/// </summary>
			/// <param name="option" type="string | string[] | OfficeExtension.LoadOption"/>
			/// <returns type="Visio.ShapeDataItem"/>
		}

		return ShapeDataItem;
	})(OfficeExtension.ClientObject);
	Visio.ShapeDataItem = ShapeDataItem;
})(Visio || (Visio = {__proto__: null}));

var Visio;
(function (Visio) {
	var ShapeDataItemCollection = (function(_super) {
		__extends(ShapeDataItemCollection, _super);
		function ShapeDataItemCollection() {
			/// <summary> Represents the ShapeDataItemCollection for a given Shape. [Api set:  1.1] </summary>
			/// <field name="context" type="Visio.RequestContext">The request context associated with this object.</field>
			/// <field name="isNull" type="Boolean">Returns a boolean value for whether the corresponding object is null. You must call "context.sync()" before reading the isNull property.</field>
			/// <field name="items" type="Array" elementType="Visio.ShapeDataItem">Gets the loaded child items in this collection.</field>
		}

		ShapeDataItemCollection.prototype.load = function(option) {
			/// <summary>
			/// Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
			/// </summary>
			/// <param name="option" type="string | string[] | OfficeExtension.LoadOption"/>
			/// <returns type="Visio.ShapeDataItemCollection"/>
		}
		ShapeDataItemCollection.prototype.getCount = function() {
			/// <summary>
			/// Gets the number of Shape Data Items. [Api set:  1.1]
			/// </summary>
			/// <returns type="OfficeExtension.ClientResult&lt;number&gt;"></returns>
			var result = new OfficeExtension.ClientResult();
			result.__proto__ = null;
			result.value = 0;
			return result;
		}
		ShapeDataItemCollection.prototype.getItem = function(key) {
			/// <summary>
			/// Gets the ShapeDataItem using its name. [Api set:  1.1]
			/// </summary>
			/// <param name="key" >Key is the name of the ShapeDataItem to be retrieved.</param>
			/// <returns type="Visio.ShapeDataItem"></returns>
		}

		return ShapeDataItemCollection;
	})(OfficeExtension.ClientObject);
	Visio.ShapeDataItemCollection = ShapeDataItemCollection;
})(Visio || (Visio = {__proto__: null}));

var Visio;
(function (Visio) {
	var Interfaces;
	(function (Interfaces) {
		var ShapeMouseEnterEventArgs = (function() {
			function ShapeMouseEnterEventArgs() {
				/// <summary> Provides information about the shape that raised the ShapeMouseEnter event. [Api set:  1.1] </summary>
				/// <field name="pageName" type="String">Gets the name of the page which has the shape object that raised the ShapeMouseEnter event. [Api set:  1.1]</field>
				/// <field name="shapeName" type="String">Gets the shape object that raised the ShapeMouseEnter event. [Api set:  1.1]</field>
			}
			return ShapeMouseEnterEventArgs;
		})();
		Interfaces.ShapeMouseEnterEventArgs.__proto__ = null;
		Interfaces.ShapeMouseEnterEventArgs = ShapeMouseEnterEventArgs;
	})(Interfaces = Visio.Interfaces || (Visio.Interfaces = { __proto__: null}));
})(Visio || (Visio = {__proto__: null}));

var Visio;
(function (Visio) {
	var Interfaces;
	(function (Interfaces) {
		var ShapeMouseLeaveEventArgs = (function() {
			function ShapeMouseLeaveEventArgs() {
				/// <summary> Provides information about the shape that raised the ShapeMouseLeave event. [Api set:  1.1] </summary>
				/// <field name="pageName" type="String">Gets the name of the page which has the shape object that raised the ShapeMouseLeave event. [Api set:  1.1]</field>
				/// <field name="shapeName" type="String">Gets the shape object that raised the ShapeMouseLeave event. [Api set:  1.1]</field>
			}
			return ShapeMouseLeaveEventArgs;
		})();
		Interfaces.ShapeMouseLeaveEventArgs.__proto__ = null;
		Interfaces.ShapeMouseLeaveEventArgs = ShapeMouseLeaveEventArgs;
	})(Interfaces = Visio.Interfaces || (Visio.Interfaces = { __proto__: null}));
})(Visio || (Visio = {__proto__: null}));

var Visio;
(function (Visio) {
	var ShapeView = (function(_super) {
		__extends(ShapeView, _super);
		function ShapeView() {
			/// <summary> Represents the ShapeView class. [Api set:  1.1] </summary>
			/// <field name="context" type="Visio.RequestContext">The request context associated with this object.</field>
			/// <field name="isNull" type="Boolean">Returns a boolean value for whether the corresponding object is null. You must call "context.sync()" before reading the isNull property.</field>
			/// <field name="highlight" type="Visio.Interfaces.Highlight">Represents the highlight around the shape. [Api set:  1.1]</field>
		}

		ShapeView.prototype.load = function(option) {
			/// <summary>
			/// Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
			/// </summary>
			/// <param name="option" type="string | string[] | OfficeExtension.LoadOption"/>
			/// <returns type="Visio.ShapeView"/>
		}

		ShapeView.prototype.set = function() {
			/// <signature>
			/// <summary>Sets multiple properties on the object at the same time, based on JSON input.</summary>
			/// <param name="properties" type="Visio.Interfaces.ShapeViewUpdateData">Properties described by the Visio.Interfaces.ShapeViewUpdateData interface.</param>
			/// <param name="options" type="string">Options of the form { throwOnReadOnly?: boolean }
			/// <br />
			/// * throwOnReadOnly: Throw an error if the passed-in property list includes read-only properties (default = true).
			/// </param>
			/// </signature>
			/// <signature>
			/// <summary>Sets multiple properties on the object at the same time, based on an existing loaded object.</summary>
			/// <param name="properties" type="ShapeView">An existing ShapeView object, with properties that have already been loaded and synced.</param>
			/// </signature>
		}
		ShapeView.prototype.addOverlay = function(OverlayType, Content, OverlayHorizontalAlignment, OverlayVerticalAlignment, Width, Height) {
			/// <summary>
			/// Adds an overlay on top of the shape. [Api set:  1.1]
			/// </summary>
			/// <param name="OverlayType" type="String">An Overlay Type -Text, Image.</param>
			/// <param name="Content" type="String">Content of Overlay.</param>
			/// <param name="OverlayHorizontalAlignment" type="String">Horizontal Alignment of Overlay - Left, Center, Right</param>
			/// <param name="OverlayVerticalAlignment" type="String">Vertical Alignment of Overlay - Top, Middle, Bottom</param>
			/// <param name="Width" type="Number">Overlay Width.</param>
			/// <param name="Height" type="Number">Overlay Height.</param>
			/// <returns type="OfficeExtension.ClientResult&lt;number&gt;"></returns>
			var result = new OfficeExtension.ClientResult();
			result.__proto__ = null;
			result.value = 0;
			return result;
		}
		ShapeView.prototype.removeOverlay = function(OverlayId) {
			/// <summary>
			/// Removes particular overlay or all overlays on the Shape. [Api set:  1.1]
			/// </summary>
			/// <param name="OverlayId" type="Number">An Overlay Id. Removes the specific overlay id from the shape.</param>
			/// <returns ></returns>
		}

		return ShapeView;
	})(OfficeExtension.ClientObject);
	Visio.ShapeView = ShapeView;
})(Visio || (Visio = {__proto__: null}));

var Visio;
(function (Visio) {
	var Interfaces;
	(function (Interfaces) {
		var ApplicationUpdateData = (function() {
			function ApplicationUpdateData() {
				/// <summary>An interface for updating data on the Application object, for use in "application.set({ ... })".</summary>
				/// <field name="showToolbars" type="Boolean">Show or Hide the standard toolbars. [Api set:  1.1]</field>;
			}
			return ApplicationUpdateData;
		})();
		Interfaces.ApplicationUpdateData.__proto__ = null;
		Interfaces.ApplicationUpdateData = ApplicationUpdateData;
	})(Interfaces = Visio.Interfaces || (Visio.Interfaces = { __proto__: null}));
})(Visio || (Visio = {__proto__: null}));

var Visio;
(function (Visio) {
	var Interfaces;
	(function (Interfaces) {
		var DocumentViewUpdateData = (function() {
			function DocumentViewUpdateData() {
				/// <summary>An interface for updating data on the DocumentView object, for use in "documentView.set({ ... })".</summary>
				/// <field name="disableHyperlinks" type="Boolean">Disable Hyperlinks. [Api set:  1.1]</field>;
				/// <field name="disablePan" type="Boolean">Disable Pan. [Api set:  1.1]</field>;
				/// <field name="disableZoom" type="Boolean">Disable Zoom. [Api set:  1.1]</field>;
				/// <field name="hideDiagramBoundary" type="Boolean">Disable Hyperlinks. [Api set:  1.1]</field>;
			}
			return DocumentViewUpdateData;
		})();
		Interfaces.DocumentViewUpdateData.__proto__ = null;
		Interfaces.DocumentViewUpdateData = DocumentViewUpdateData;
	})(Interfaces = Visio.Interfaces || (Visio.Interfaces = { __proto__: null}));
})(Visio || (Visio = {__proto__: null}));

var Visio;
(function (Visio) {
	var Interfaces;
	(function (Interfaces) {
		var PageViewUpdateData = (function() {
			function PageViewUpdateData() {
				/// <summary>An interface for updating data on the PageView object, for use in "pageView.set({ ... })".</summary>
				/// <field name="zoom" type="Number">Get/Set Page&apos;s Zoom level. The value can be between 10 and 400 and denotes the percentage of zoom. [Api set:  1.1]</field>;
			}
			return PageViewUpdateData;
		})();
		Interfaces.PageViewUpdateData.__proto__ = null;
		Interfaces.PageViewUpdateData = PageViewUpdateData;
	})(Interfaces = Visio.Interfaces || (Visio.Interfaces = { __proto__: null}));
})(Visio || (Visio = {__proto__: null}));

var Visio;
(function (Visio) {
	var Interfaces;
	(function (Interfaces) {
		var ShapeUpdateData = (function() {
			function ShapeUpdateData() {
				/// <summary>An interface for updating data on the Shape object, for use in "shape.set({ ... })".</summary>
				/// <field name="select" type="Boolean">Returns true, if shape is selected. User can set true to select the shape explicitly. [Api set:  1.1]</field>;
			}
			return ShapeUpdateData;
		})();
		Interfaces.ShapeUpdateData.__proto__ = null;
		Interfaces.ShapeUpdateData = ShapeUpdateData;
	})(Interfaces = Visio.Interfaces || (Visio.Interfaces = { __proto__: null}));
})(Visio || (Visio = {__proto__: null}));

var Visio;
(function (Visio) {
	var Interfaces;
	(function (Interfaces) {
		var ShapeViewUpdateData = (function() {
			function ShapeViewUpdateData() {
				/// <summary>An interface for updating data on the ShapeView object, for use in "shapeView.set({ ... })".</summary>
				/// <field name="highlight" type="Visio.Interfaces.Highlight">Represents the highlight around the shape. [Api set:  1.1]</field>;
			}
			return ShapeViewUpdateData;
		})();
		Interfaces.ShapeViewUpdateData.__proto__ = null;
		Interfaces.ShapeViewUpdateData = ShapeViewUpdateData;
	})(Interfaces = Visio.Interfaces || (Visio.Interfaces = { __proto__: null}));
})(Visio || (Visio = {__proto__: null}));

var Visio;
(function (Visio) {
	var Interfaces;
	(function (Interfaces) {
		var CommentUpdateData = (function() {
			function CommentUpdateData() {
				/// <summary>An interface for updating data on the Comment object, for use in "comment.set({ ... })".</summary>
				/// <field name="author" type="String">A string that specifies the label of the shape data item. [Api set:  1.1]</field>;
				/// <field name="date" type="String">A string that specifies the format of the shape data item. [Api set:  1.1]</field>;
				/// <field name="text" type="String">A string that specifies the value of the shape data item. [Api set:  1.1]</field>;
			}
			return CommentUpdateData;
		})();
		Interfaces.CommentUpdateData.__proto__ = null;
		Interfaces.CommentUpdateData = CommentUpdateData;
	})(Interfaces = Visio.Interfaces || (Visio.Interfaces = { __proto__: null}));
})(Visio || (Visio = {__proto__: null}));
var Visio;
(function (Visio) {
	var RequestContext = (function (_super) {
		__extends(RequestContext, _super);
		function RequestContext() {
			/// <summary>
			/// The RequestContext object facilitates requests to the Visio application. Since the Office add-in and the Visio application run in two different processes, the request context is required to get access to the Visio object model from the add-in.
			/// </summary>
			/// <field name="document" type="Visio.Document">Root object for interacting with the document</field>
			_super.call(this, null);
		}
		return RequestContext;
	})(OfficeExtension.ClientRequestContext);
	Visio.RequestContext = RequestContext;

	Visio.run = function (batch) {
		/// <signature>
		/// <summary>
		/// Executes a batch script that performs actions on the Visio object model, using a new RequestContext. When the promise is resolved, any tracked objects that were automatically allocated during execution will be released.
		/// </summary>
		/// <param name="batch" type="function(context) { ... }">
		/// A function that takes in a RequestContext and returns a promise (typically, just the result of "context.sync()").
		/// <br />
		/// The context parameter facilitates requests to the Visio application. Since the Office add-in and the Visio application run in two different processes, the RequestContext is required to get access to the Visio object model from the add-in.
		/// </param>
		/// </signature>
		/// <signature>
		/// <summary>
		/// Executes a batch script that performs actions on the Visio object model, using the RequestContext of a previously-created API object. When the promise is resolved, any tracked objects that were automatically allocated during execution will be released.
		/// </summary>
		/// <param name="object" type="OfficeExtension.ClientObject">
		/// A previously-created API object. The batch will use the same RequestContext as the passed-in object, which means that any changes applied to the object will be picked up by "context.sync()".
		/// </param>
		/// <param name="batch" type="function(context) { ... }">
		/// A function that takes in a RequestContext and returns a promise (typically, just the result of "context.sync()").
		/// <br />
		/// The context parameter facilitates requests to the Visio application. Since the Office add-in and the Visio application run in two different processes, the RequestContext is required to get access to the Visio object model from the add-in.
		/// </param>
		/// </signature>
		/// <signature>
		/// <summary>
		/// Executes a batch script that performs actions on the Visio object model, using the RequestContext of a previously-created API object. When the promise is resolved, any tracked objects that were automatically allocated during execution will be released.
		/// </summary>
		/// <param name="objects" type="Array&lt;OfficeExtension.ClientObject&gt;">
		/// An array of previously-created API objects. The array will be validated to make sure that all of the objects share the same context. The batch will use this shared RequestContext, which means that any changes applied to these objects will be picked up by "context.sync()".
		/// </param>
		/// <param name="batch" type="function(context) { ... }">
		/// A function that takes in a RequestContext and returns a promise (typically, just the result of "context.sync()").
		/// <br />
		/// The context parameter facilitates requests to the Visio application. Since the Office add-in and the Visio application run in two different processes, the RequestContext is required to get access to the Visio object model from the add-in.
		/// </param>
		/// </signature>
		arguments[arguments.length - 1](new Visio.RequestContext());
		return new OfficeExtension.Promise();
	}
})(Visio || (Visio = {__proto__: null}));
Visio.__proto__ = null;

