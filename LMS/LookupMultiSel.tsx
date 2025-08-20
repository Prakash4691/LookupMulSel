import * as React from "react";
import {
  Dropdown,
  IDropdownOption,
  IDropdownStyles,
  DropdownMenuItemType,
} from "@fluentui/react/lib/Dropdown";
import { Stack, IStackTokens } from "@fluentui/react/lib/Stack";
import { IInputs } from "./generated/ManifestTypes";
import { IconButton, IButtonStyles } from "@fluentui/react/lib/Button";
import { Icon } from "@fluentui/react/lib/Icon";
import { SearchBox } from "@fluentui/react/lib/SearchBox";
import { FontSizes, ISearchBoxStyles, ITheme } from "@fluentui/react";
import { Spinner, SpinnerSize } from "@fluentui/react/lib/Spinner";
import { MessageBar, MessageBarType } from "@fluentui/react/lib/MessageBar";
import { Associate, DisAssociate } from "./WebApiOperations";

const dropdownStyles: Partial<IDropdownStyles> = {
  dropdown: {
    width: "100%",
    selectors: {
      // Remove strong outer focus ring; keep outline off to rely on internal element focus visuals
      "&:focus": {
        borderColor: "#e1dfdd", // keep neutral border instead of blue highlight
        boxShadow: "none",
        outline: "none",
      },
      "&:hover": {
        borderColor: "#d2d0ce",
      },
    },
  },
  root: {
    width: "100%",
  },
  callout: {
    maxHeight: "200px",
    borderRadius: "2px",
    boxShadow:
      "0 1.6px 3.6px 0 rgba(0, 0, 0, 0.132), 0 0.3px 0.9px 0 rgba(0, 0, 0, 0.108)",
    border: "1px solid #e1dfdd",
  },
  dropdownItemsWrapper: {
    maxHeight: "inherit",
  },
  title: {
    borderRadius: "2px",
    minHeight: "32px",
    padding: "8px 40px 8px 12px", // Updated padding to match CSS
    fontSize: "14px",
    lineHeight: "20px",
    backgroundColor: "#ffffff",
    selectors: {
      "&:hover": {
        borderColor: "#d2d0ce",
      },
    },
  },
  caretDownWrapper: {
    color: "#605e5c",
    fontSize: "16px",
    selectors: {
      "&:hover": {
        color: "#0078d4",
      },
    },
  },
};

const searchBoxStyles: Partial<ISearchBoxStyles> = {
  clearButton: { display: "none" },
  root: {
    margin: "8px 12px",
    borderRadius: "2px",
  },
  field: {
    fontSize: "14px",
    padding: "8px 12px",
    border: "1px solid #d2d0ce",
    borderRadius: "2px",
  },
};

const stackTokens: IStackTokens = {
  childrenGap: 4,
  padding: 0,
};

const buttonStyles: IButtonStyles = {
  icon: {
    fontSize: "12px",
    color: "#666666",
  },
  root: {
    height: "24px",
    width: "24px",
    minWidth: "24px",
  },
};

export interface ILookupMultiSel {
  onChange: (selectedValues: string[]) => void;
  initialValues: string[];
  context: ComponentFramework.Context<IInputs>;
  relatedEntityType: string;
  relatedPrimaryColumns: string[];
  primaryEntityType: string;
  relationshipName: string;
  primaryEntityId: string;
  disabled: boolean;
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
    disabled,
  } = props;
  const [selectedValues, setSelectedValues] = React.useState<string[]>([]);
  const [userOptions, setUserOptions] = React.useState<IDropdownOption[]>([]);
  const [isLoading, setIsLoading] = React.useState<boolean>(true);
  const [error, setError] = React.useState<string>("");
  const onChangeTriggered = React.useRef(false);
  const [searchText, setSearchText] = React.useState<string>("");
  const [debouncedSearchText, setDebouncedSearchText] =
    React.useState<string>("");
  const searchTimeoutRef = React.useRef<any>();
  // Derived dropdown state - smart filtering that preserves selected options
  const selectedOptions = userOptions.filter((opt) =>
    selectedValues.includes(opt.key as string)
  );
  const unselectedOptions = userOptions.filter(
    (opt) => !selectedValues.includes(opt.key as string)
  );
  const filteredUnselectedOptions = unselectedOptions.filter(
    (opt) =>
      opt.text
        .toLocaleLowerCase()
        .indexOf(debouncedSearchText.toLocaleLowerCase()) > -1
  );
  const filteredOptions = [...selectedOptions, ...filteredUnselectedOptions];
  const hasSearchText = debouncedSearchText.trim().length > 0;
  const hasResults = filteredUnselectedOptions.length > 0 || selectedOptions.length > 0;
  const noResults = hasSearchText && filteredUnselectedOptions.length === 0 && selectedOptions.length === 0;

  // Adjust dropdown styles to remove scrolling when "no results" message shown
  const adjustedDropdownStyles = React.useMemo(() => {
    if (noResults) {
      return {
        ...dropdownStyles,
        callout: {
          ...((dropdownStyles.callout as object) || {}),
          maxHeight: "none",
        },
        dropdownItemsWrapper: {
          maxHeight: "none",
        },
      } as Partial<IDropdownStyles>;
    }
    return dropdownStyles;
  }, [noResults]);

  /**
   * Debounce search text to improve performance
   */
  React.useEffect(() => {
    if (searchTimeoutRef.current) {
      clearTimeout(searchTimeoutRef.current);
    }

    searchTimeoutRef.current = setTimeout(() => {
      setDebouncedSearchText(searchText);
    }, 300);

    return () => {
      if (searchTimeoutRef.current) {
        clearTimeout(searchTimeoutRef.current);
      }
    };
  }, [searchText]);

  /**
   * Sync selected values with incoming initialValues (handles re-publish / control reload)
   */
  React.useEffect(() => {
    const areEqual = (a: string[] = [], b: string[] = []) => {
      if (a.length !== b.length) return false;
      const as = [...a].sort();
      const bs = [...b].sort();
      for (let i = 0; i < as.length; i++) if (as[i] !== bs[i]) return false;
      return true;
    };
    if (!areEqual(selectedValues, initialValues)) {
      setSelectedValues(initialValues ?? []);
    }
  }, [initialValues]);

  /**
   * Retrieves entity records using webapi and maintain using state
   */
  React.useEffect(() => {
    let userOptionsList: IDropdownOption[] = [];
    setIsLoading(true);
    setError("");

    context.webAPI
      .retrieveMultipleRecords(relatedEntityType)
      .then((response) => {
        response.entities.forEach((element) => {
          userOptionsList.push({
            key: element[relatedPrimaryColumns[0]],
            text: element[relatedPrimaryColumns[1]],
            data: { value: element[relatedPrimaryColumns[0]] },
          });
        });
        setUserOptions(userOptionsList);
        setIsLoading(false);
      })
      .catch((error) => {
        console.error("Error retrieving records:", error);
        setError("Failed to load records. Please try again.");
        setIsLoading(false);
        context.navigation.openAlertDialog(error);
      });
  }, []);

  /**
   * Trigger onchange to update the property
   */
  React.useEffect(() => {
    if (onChangeTriggered.current) onChange(selectedValues);
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
    if (option?.itemType === DropdownMenuItemType.Header) {
      if (option.key === "FilterHeader") {
        return (
          <SearchBox
            onChange={(ev, newValue?: string) => setSearchText(newValue || "")}
            underlined={false}
            placeholder="Search options..."
            autoFocus={true}
            styles={searchBoxStyles}
            iconProps={{ iconName: "Search" }}
            clearButtonProps={{
              ariaLabel: "Clear search",
            }}
            ariaLabel="Search dropdown options"
          />
        );
      } else if (option.key === "no-results") {
        return (
          <div className="no-results-message">
            <Icon iconName="Search" className="no-results-icon" />
            {option.text}
          </div>
        );
      }
    }

    return <div className="dropdown-option-item">{option?.text}</div>;
  };

  /**
   * Render custom title with selected items as chips/tags
   * @param options Selected option from dropdown
   * @returns
   */
  const onRenderTitle = (options: any) => {
    const selectedList: IDropdownOption[] = options;

    if (selectedList.length === 0) {
      return <span className="placeholder-text">Look for records</span>;
    }

    return (
      <Stack
        horizontal
        wrap
        tokens={stackTokens}
        className="selected-items-container"
      >
        {selectedList.map((element, index) => (
          <div key={element.key} className="selected-item-chip">
            <span title={element.text} className="selected-item-text">
              {element.text}
            </span>
            <IconButton
              iconProps={{ iconName: "Cancel" }}
              title={`Remove ${element.text}`}
              onClick={(ev) => onChangeDropDownOrOnIconClick(ev, element, 1)}
              className="remove-chip-button"
            />
          </div>
        ))}
      </Stack>
    );
  };

  return (
    <>
      {error && (
        <div className="error-container">
          <MessageBar
            messageBarType={MessageBarType.error}
            isMultiline={false}
            onDismiss={() => setError("")}
          >
            {error}
          </MessageBar>
        </div>
      )}

      {isLoading ? (
        <div className="loading-container">
          <Spinner size={SpinnerSize.medium} label="Loading records..." />
        </div>
      ) : (
        <Dropdown
          {...userOptions}
          options={((): IDropdownOption[] => {
            const options: IDropdownOption[] = [
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
            ];

            if (noResults) {
              options.push({
                key: "no-results",
                text: `No results found for "${debouncedSearchText}"`,
                itemType: DropdownMenuItemType.Header,
              });
            } else {
              options.push(...filteredOptions);
            }
            return options;
          })()}
          styles={adjustedDropdownStyles}
          multiSelect={true}
          onChange={onChangeDropDownOrOnIconClick}
          selectedKeys={selectedValues}
          calloutProps={{
            directionalHintFixed: true,
            isBeakVisible: false,
          }}
          onRenderTitle={onRenderTitle}
          id="MainDropDown"
          placeholder="Look for records"
          onRenderCaretDown={onRenderCaretDown}
          onRenderOption={onRenderOption}
          onDismiss={() => {
            setSearchText("");
            setDebouncedSearchText("");
          }}
          disabled={disabled}
          ariaLabel="Multi-select lookup field"
          data-testid="lookup-multiselect-dropdown"
        />
      )}
    </>
  );
});

LookupMultiSel.displayName = "LookupMultiSel";
