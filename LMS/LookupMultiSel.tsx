import * as React from "react";
import {
  Dropdown,
  IDropdownOption,
  IDropdownStyles,
  DropdownMenuItemType,
} from "@fluentui/react/lib/Dropdown";
import { IInputs } from "./generated/ManifestTypes";
import {
  IconButton,
  IButtonStyles,
  PrimaryButton,
} from "@fluentui/react/lib/Button";
import { Icon } from "@fluentui/react/lib/Icon";
import { SearchBox } from "@fluentui/react/lib/SearchBox";
import { ISearchBoxStyles } from "@fluentui/react/lib/SearchBox";
import { Associate, DisAssociate } from "./WebApiOperations";

const dropdownStyles: Partial<IDropdownStyles> = {
  dropdown: { width: 500, height: "auto" },
};

const searchBoxStyles: Partial<ISearchBoxStyles> = {
  clearButton: { display: "none" },
};
const buttonStyles: IButtonStyles = { icon: { fontSize: "11px" } };

/**
 * Interface for LookupMultiSel component props
 */
export interface ILookupMultiSel {
  onChange: (selectedValues: string[]) => void; // Callback when selection changes
  initialValues: string[]; // Initially selected values
  context: ComponentFramework.Context<IInputs>; // Component context
  relatedEntityType: string; // Type of related entity
  relatedPrimaryColumns: string[]; // Primary columns of related entity
  primaryEntityType: string; // Type of primary entity
  relationshipName: string; // Name of the relationship
  primaryEntityId: string; // ID of the primary entity
  disabled: boolean; // Whether the control is disabled
  maxPageSize: number; // Maximum number of records per page
}

/**
 * A multi-select lookup component that allows users to search and select multiple related records
 * Features include:
 * - Search functionality
 * - Pagination support
 * - Associate/Disassociate operations
 * - Custom rendering of selected items with remove buttons
 */
export const LookupMultiSel = React.memo((props: ILookupMultiSel) => {
  const {
    onChange,
    initialValues,
    context,
    relatedEntityType,
    relatedPrimaryColumns,
    primaryEntityType,
    relationshipName,
    primaryEntityId,
    disabled,
  } = props;

  const [selectedValues, setSelectedValues] = React.useState<string[]>([]);
  const [filteredRelatedRecords, setFilteredRelatedRecords] = React.useState<
    IDropdownOption[]
  >([]);
  const [fullRelatedRecords, setFullRelatedRecords] = React.useState<
    IDropdownOption[]
  >([]);
  const [selectedOptions, setSelectedOptions] = React.useState<
    IDropdownOption[]
  >([]);
  const onChangeTriggered = React.useRef(false);
  const [searchText, setSearchText] = React.useState<string>("");
  const [IsLoadMoreButtonDisabled, setLoadMoreButtonDisabled] =
    React.useState(true);
  const [IsPreviousButtonDisabled, setPreviousButtonDisabled] =
    React.useState(true);
  const [pageNumber, setPageNumber] = React.useState(1);
  const [pagingCookieValue, setPagingCookie] = React.useState<string>("");
  const [previousPageURL, setPreviousPageURL] = React.useState<string>("");
  const [isLoadMoreButtonClicked, setIsLoadMoreButtonClicked] =
    React.useState<boolean>(false);
  const [ispreviousButtonClicked, setIspreviousButtonClicked] =
    React.useState<boolean>(false);

  /**
   * Fetches records from the related entity with optional paging
   * @param isPrevious - Boolean flag indicating if fetching previous page
   */
  const fetchRecords = async (isPrevious?: boolean) => {
    let userOptionsList: IDropdownOption[] = [];
    let query = setStateValuesAndRetrieveQuery(isPrevious);

    try {
      let response = await context.webAPI.retrieveMultipleRecords(
        relatedEntityType,
        query,
        props.maxPageSize
      );

      response.entities.map((element) => {
        userOptionsList.push({
          key: element[relatedPrimaryColumns[0]],
          text: element[relatedPrimaryColumns[1]],
          data: { value: element[relatedPrimaryColumns[0]] },
        });
      });

      /*if (updatedCookie) {
        setPagingCookie(updatedCookie);
        setLoadMoreButtonDisabled(false);
      } else */
      if (response.nextLink) {
        const nextLink = response.nextLink;
        const skipTokenParam = "$skiptoken=";
        const skipTokenIndex = nextLink.indexOf(skipTokenParam);
        if (skipTokenIndex !== -1) {
          const nextPageCookie = nextLink.substring(
            skipTokenIndex + skipTokenParam.length
          );
          setPagingCookie(nextPageCookie);
          setLoadMoreButtonDisabled(false);
        }
      } else {
        setLoadMoreButtonDisabled(true);
      }

      //setUserOptions([...userOptions, ...userOptionsList]);
      setFilteredRelatedRecords(userOptionsList);
    } catch (error) {
      context.navigation.openAlertDialog({ text: (error as Error).message });
    }
  };

  /**
   * Fetches all records from the related entity without paging
   * Used to populate the full list of related records and initialize selected options
   */
  const fetchAllRecords = async () => {
    let userOptionsList: IDropdownOption[] = [];
    let response = await context.webAPI.retrieveMultipleRecords(
      relatedEntityType,
      `?$select=${relatedPrimaryColumns.join(",")}`
    );

    response.entities.map((element) => {
      userOptionsList.push({
        key: element[relatedPrimaryColumns[0]],
        text: element[relatedPrimaryColumns[1]],
        data: { value: element[relatedPrimaryColumns[0]] },
      });
    });

    setSelectedOptions(
      userOptionsList.filter((opt) =>
        props.initialValues.includes(opt.key as string)
      )
    );

    setFullRelatedRecords(userOptionsList);
  };

  /**
   * Prepares the query string and updates paging-related state values
   * @param isPrevious - Boolean flag indicating if fetching previous page
   * @returns Query string for the WebAPI call
   */
  const setStateValuesAndRetrieveQuery = (isPrevious?: boolean) => {
    setLoadMoreButtonDisabled(true);
    if (isLoadMoreButtonClicked) setIsLoadMoreButtonClicked(false);
    if (ispreviousButtonClicked) setIspreviousButtonClicked(false);
    let query = `?$select=${relatedPrimaryColumns.join(",")}`;

    if (
      isPrevious &&
      pageNumber === 1 &&
      previousPageURL.includes("&$skiptoken=")
    ) {
      query = previousPageURL.replace(/&\$skiptoken=.*$/, "");
    } else if (isPrevious) {
      query = previousPageURL;
    } else if (pagingCookieValue) {
      query += `&$skiptoken=${pagingCookieValue}`;
    }

    if (pageNumber === 1) {
      setPreviousPageURL("");
    } else if (pageNumber === 2) {
      setPreviousPageURL(query.replace(`&$skiptoken=${pagingCookieValue}`, ""));
    } else if (pageNumber > 2) {
      let updatedCookie = updatePageNumberInCookie(
        pagingCookieValue,
        pageNumber - 1
      );
      setPreviousPageURL(
        `?$select=${relatedPrimaryColumns.join(
          ","
        )}&$skiptoken=${updatedCookie}`
      );
    }

    return query;
  };

  /**
   * Updates the page number in the paging cookie string
   * @param cookie - The current paging cookie string
   * @param newPageNumber - The new page number to set
   * @returns Updated cookie string with new page number
   */
  const updatePageNumberInCookie = (cookie: string, newPageNumber: number) => {
    const pageNumberRegex = /pagenumber=%22(\d+)%22/;
    return cookie.replace(pageNumberRegex, `pagenumber=%22${newPageNumber}%22`);
  };

  /**
   * Gets selected values from props and maintain using state
   * Retrieves entity records using webapi and maintain using state
   */
  React.useEffect(() => {
    fetchRecords();
    setSelectedValues(initialValues);
    fetchAllRecords();
  }, []);

  /**
   * Trigger onchange to update the property
   */
  React.useEffect(() => {
    let selectedValueOptions: string[] = [];
    selectedValueOptions.push(selectedValues.toString());

    if (selectedValues)
      setSelectedOptions(
        fullRelatedRecords.filter((opt) =>
          selectedValues.includes(opt.key as string)
        )
      );

    if (onChangeTriggered.current) onChange(selectedValueOptions);
  }, [selectedValues]);

  /**
   * Handles changes in dropdown selection or cancel icon clicks
   * @param ev - Event object
   * @param option - The dropdown option that was selected/deselected
   * @param eventId - Optional identifier to distinguish between dropdown and cancel icon events (1 for cancel icon)
   */
  const onChangeDropDownOrOnIconClick = (
    ev: unknown,
    option?: IDropdownOption,
    eventId?: number
  ) => {
    if (!option) return;

    if (eventId === 1) {
      (ev as React.MouseEvent<HTMLButtonElement>).stopPropagation();
    }

    onChangeTriggered.current = true;
    setSelectedValues(
      option.selected
        ? [...selectedValues, option.key as string]
        : selectedValues.filter((key) => key !== option.key)
    );

    if (option?.selected)
      Associate(
        context,
        option.key,
        primaryEntityType,
        relatedEntityType,
        relationshipName,
        primaryEntityId
      );
    else
      DisAssociate(
        context,
        option?.key!,
        primaryEntityType,
        relationshipName,
        primaryEntityId
      );
  };

  /**
   * Renders the search icon in place of the default dropdown caret
   * @returns Icon component
   */
  const onRenderCaretDown = () => {
    return <Icon iconName="Search"></Icon>;
  };

  /**
   * Custom renderer for dropdown options
   * Renders either a search box for the header or the option text
   * @param option - The dropdown option to render
   * @returns Rendered component
   */
  const onRenderOption = (option?: IDropdownOption) => {
    if (!option) return null;
    if (
      option.itemType === DropdownMenuItemType.Header &&
      option.key === "FilterHeader"
    ) {
      return (
        <SearchBox
          onChange={(ev, newValue?: string) => setSearchText(newValue!)}
          underlined={true}
          placeholder="Search options"
          autoFocus={true}
          styles={searchBoxStyles}
        />
      );
    }
    return <>{option.text}</>;
  };

  /**
   * Custom renderer for the dropdown title
   * Renders selected options with cancel icons
   * @returns Rendered component with selected options
   */
  const onRenderTitle = () => {
    return (
      <div>
        {selectedOptions?.map((option) => (
          <span key={option.key}>
            {option.text}
            <IconButton
              iconProps={{ iconName: "Cancel" }}
              title={option.text}
              onClick={(ev) => onChangeDropDownOrOnIconClick(ev, option, 1)}
              className="IconButtonClass"
              styles={buttonStyles}
            />
          </span>
        ))}
      </div>
    );
  };

  React.useEffect(() => {
    if (pageNumber === 1) setPreviousButtonDisabled(true);
  }, [pageNumber]);

  React.useMemo(() => {
    if (isLoadMoreButtonClicked) fetchRecords(false);
  }, [isLoadMoreButtonClicked]);

  React.useMemo(() => {
    if (ispreviousButtonClicked) fetchRecords(true);
  }, [ispreviousButtonClicked]);

  /**
   * Loads the next page of records
   * Updates page number and triggers record fetch
   */
  const loadMore = () => {
    setPageNumber((prevPageNumber) => prevPageNumber + 1);
    setPreviousButtonDisabled(false);
    setIsLoadMoreButtonClicked(true);
  };

  /**
   * Loads the previous page of records
   * Updates page number and triggers record fetch
   */
  const loadPrevious = () => {
    setPageNumber((prevPageNumber) => prevPageNumber - 1);
    setIspreviousButtonClicked(true);
    /*const match = pagingCookieValue.match(/pagenumber=%22(\d+)%22/);
    const previousPageNumber = match ? parseInt(match[1], 10) - 1 : 1;
    if (pageNumber > 2) fetchRecords(true, previousPageNumber);
    else {
      setPreviousButtonDisabled(true);
      fetchRecords(false);
    }*/
  };

  const mergedOptions = [
    {
      key: "FilterHeader",
      text: "-",
      itemType: DropdownMenuItemType.Header,
    },
    {
      key: "divider_filterHeader",
      text: "-",
      itemType: DropdownMenuItemType.Divider,
    },
    ...filteredRelatedRecords.filter((opt) =>
      opt.text.toLocaleLowerCase().includes(searchText.toLocaleLowerCase())
    ),
  ];

  return (
    <>
      <Dropdown
        {...filteredRelatedRecords}
        options={mergedOptions}
        styles={dropdownStyles}
        multiSelect={true}
        onChange={onChangeDropDownOrOnIconClick}
        selectedKeys={selectedValues}
        // Append a footer by customizing the list rendering
        onRenderList={(props, defaultRender) => (
          <div>
            {defaultRender && defaultRender(props)}
            <div
              style={{
                display: "flex",
                justifyContent: "space-between",
                padding: "8px",
                borderTop: "1px solid #ccc",
              }}
            >
              <Icon
                iconName="ChevronUp"
                onClick={(ev) => {
                  ev.stopPropagation();
                  if (!IsPreviousButtonDisabled) loadPrevious();
                }}
                style={{
                  cursor: IsPreviousButtonDisabled ? "default" : "pointer",
                }}
              />
              <Icon
                iconName="ChevronDown"
                onClick={(ev) => {
                  ev.stopPropagation();
                  if (!IsLoadMoreButtonDisabled) loadMore();
                }}
                style={{
                  cursor: IsLoadMoreButtonDisabled ? "default" : "pointer",
                }}
              />
            </div>
          </div>
        )}
        onRenderTitle={onRenderTitle}
        id="MainDropDown"
        placeholder="Look for records"
        onRenderCaretDown={onRenderCaretDown}
        onRenderOption={onRenderOption}
        onDismiss={() => setSearchText("")}
        disabled={disabled}
      />
      {/* <PrimaryButton
        text="Previous"
        onClick={loadPrevious}
        disabled={IsPreviousButtonDisabled}
      />
      <PrimaryButton
        text="Load More.."
        onClick={loadMore}
        disabled={IsLoadMoreButtonDisabled}
      /> */}
    </>
  );
});

LookupMultiSel.displayName = "LookupMultiSel";
