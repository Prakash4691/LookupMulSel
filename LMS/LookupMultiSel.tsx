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

export interface ILookupMultiSel {
  onChange: (selectedValues: string[]) => void;
  initialValues: string[];
  context: ComponentFramework.Context<IInputs>;
  relatedEntityType: string;
  relatedPrimaryColumns: string[];
  primaryEntityType: string;
  relationshipName: string;
  primaryEntityId: string;
  isEnabled: boolean;
  maxPageSize: number;
}

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
    isEnabled,
  } = props;

  const [selectedValues, setSelectedValues] = React.useState<string[]>([]);
  const [userOptions, setUserOptions] = React.useState<IDropdownOption[]>([]);
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

  const fetchRecords = async (isPrevious?: boolean) =>
    //isPrevious?: boolean,
    //previousPageNumber?: number
    {
      let userOptionsList: IDropdownOption[] = [];
      let query = setStateValuesAndRetrieveQuery(isPrevious);
      /*let updatedCookie: string = "";
    if (pagingCookieValue && isPrevious && previousPageNumber! > 1) {
      updatedCookie = updatePageNumberInCookie(
        pagingCookieValue,
        previousPageNumber!
      );
      query += `&$skiptoken=${updatedCookie}`;
    } else if (pagingCookieValue && pageNumber > 1) {
      query += `&$skiptoken=${pagingCookieValue}`;
    }*/
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
        setUserOptions(userOptionsList);
      } catch (error) {
        context.navigation.openAlertDialog({ text: (error as Error).message });
      }
    };

  /**
   * Set state values and retrieve query
   * @param isPrevious
   * @returns
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
   *
   * @param cookie
   * @param newPageNumber
   * @returns
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
  }, []);

  /**
   * Trigger onchange to update the property
   */
  React.useEffect(() => {
    let selectedValueOptions: string[] = [];
    selectedValueOptions.push(selectedValues.toString());
    if (onChangeTriggered.current) onChange(selectedValueOptions);
  }, [selectedValues]);

  /**
   * Triggers on change of dropdown
   * @param ev Event of the dropdown
   * @param option Selected option from dropdown
   * @param eventId Event to identify is it for dropdown or cancel icon
   */
  const onChangeDropDownOrOnIconClick = (
    ev: unknown,
    option?: IDropdownOption,
    eventId?: number
  ) => {
    if (eventId === 1) {
      let iconEvent = ev as React.MouseEvent<HTMLButtonElement>;
      iconEvent.stopPropagation();
    }

    if (option) {
      onChangeTriggered.current = true;
      setSelectedValues(
        option.selected
          ? [...selectedValues, option.key as string]
          : selectedValues.filter((key) => key != option.key)
      );
    }

    if (option?.selected)
      Associate(
        context,
        option.key,
        primaryEntityType,
        relatedEntityType,
        relationshipName,
        primaryEntityId
      );
    else if (!option?.selected)
      DisAssociate(
        context,
        option?.key!,
        primaryEntityType,
        relationshipName,
        primaryEntityId
      );
  };

  /**
   *Render icon of the dropdown to search
   * @returns Icon
   */
  const onRenderCaretDown = () => {
    return <Icon iconName="Search"></Icon>;
  };

  /**
   * Render drop down item event
   * @param option Drop down item
   * @returns
   */
  const onRenderOption = (option?: IDropdownOption) => {
    return option?.itemType === DropdownMenuItemType.Header &&
      option.key === "FilterHeader" ? (
      <SearchBox
        onChange={(ev, newValue?: string) => setSearchText(newValue!)}
        underlined={true}
        placeholder="Search options"
        autoFocus={true}
        styles={searchBoxStyles}
      ></SearchBox>
    ) : (
      <>{option?.text}</>
    );
  };

  /**
   * Render custom title
   * @param options Selected option from dropdown
   * @returns
   */
  const onRenderTitle = (options: any) => {
    let option: any[] = [];
    let selectedList: IDropdownOption[] = options;
    selectedList.forEach((element) => {
      option.push(
        <span>
          {element.text}
          <IconButton
            iconProps={{ iconName: "Cancel" }}
            title={element.text}
            onClick={(ev) => onChangeDropDownOrOnIconClick(ev, element, 1)}
            className="IconButtonClass"
            styles={buttonStyles}
          ></IconButton>
        </span>
      );
    });
    return <div>{option}</div>;
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

  const loadMore = () => {
    setPageNumber((prevPageNumber) => prevPageNumber + 1);
    setPreviousButtonDisabled(false);
    setIsLoadMoreButtonClicked(true);
  };

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

  return (
    <>
      <Dropdown
        {...userOptions}
        options={[
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
          ...userOptions.filter(
            (opt) =>
              opt.text
                .toLocaleLowerCase()
                .indexOf(searchText.toLocaleLowerCase()) > -1
          ),
        ]}
        styles={dropdownStyles}
        multiSelect={true}
        onChange={onChangeDropDownOrOnIconClick}
        selectedKeys={selectedValues}
        calloutProps={{ directionalHintFixed: true }}
        onRenderTitle={onRenderTitle}
        id="MainDropDown"
        placeholder="Look for records"
        onRenderCaretDown={onRenderCaretDown}
        onRenderOption={onRenderOption}
        onDismiss={() => setSearchText("")}
        disabled={isEnabled}
      />
      <PrimaryButton
        text="Previous"
        onClick={loadPrevious}
        disabled={IsPreviousButtonDisabled}
      />
      <PrimaryButton
        text="Load More.."
        onClick={loadMore}
        disabled={IsLoadMoreButtonDisabled}
      />
    </>
  );
});

LookupMultiSel.displayName = "LookupMultiSel";
