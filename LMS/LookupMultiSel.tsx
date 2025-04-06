import * as React from "react";
import {
  Dropdown,
  IDropdownOption,
  IDropdownStyles,
  DropdownMenuItemType,
} from "@fluentui/react/lib/Dropdown";
//import { Stack, IStackTokens } from "@fluentui/react/lib/Stack";
import { IInputs } from "./generated/ManifestTypes";
import { IconButton, IButtonStyles } from "@fluentui/react/lib/Button";
import { Icon } from "@fluentui/react/lib/Icon";
import { SearchBox } from "@fluentui/react/lib/SearchBox";
import { FontSizes, ISearchBoxStyles, ITheme } from "@fluentui/react";
//import { Sticky } from "@fluentui/react";
import { Associate, DisAssociate } from "./WebApiOperations";

const dropdownStyles: Partial<IDropdownStyles> = {
  dropdown: {
    width: "100%",
    selectors: {
      "&:focus": {
        borderColor: "#0078d4",
        boxShadow: "0 0 5px rgba(0, 120, 212, 0.5)",
      },
    },
  },
  root: {
    width: "100%",
  },
  callout: {
    maxHeight: "50vh",
    overflowY: "auto",
  },
  dropdownItemsWrapper: {
    maxHeight: "inherit",
  },
  title: {
    borderColor: "#666666",
    selectors: {
      "&:hover": {
        borderColor: "#333333",
      },
    },
  },
};

const searchBoxStyles: Partial<ISearchBoxStyles> = {
  clearButton: { display: "none" },
};
//const stackTokens: IStackTokens = { childrenGap: 10 };
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
  const onChangeTriggered = React.useRef(false);
  const [searchText, setSearchText] = React.useState<string>("");

  /**
   * Gets selected values from props and maintain using state
   */
  React.useEffect(() => {
    setSelectedValues(initialValues);
  }, []);

  /**
   * Retrieves entity records using webapi and maintain using state
   */
  React.useEffect(() => {
    let userOptionsList: IDropdownOption[] = [];
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
      })
      .catch((error) => {
        context.navigation.openAlertDialog(error);
      });
    /* let userOptionsList = RetrieveMultiple(context, entityType, entityColumns);
    setUserOptions(userOptionsList); */
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
    //let url: string = `main.aspx?pagetype=entityrecord&etn=${entityType}&id=`;
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

  return (
    <>
      {/* <Stack horizontal tokens={stackTokens}> */}
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
        //dropdownWidth="auto"
        id="MainDropDown"
        placeholder="Look for records"
        onRenderCaretDown={onRenderCaretDown}
        onRenderOption={onRenderOption}
        onDismiss={() => setSearchText("")}
        disabled={disabled}
      />
      {/* </Stack> */}
    </>
  );
});

LookupMultiSel.displayName = "LookupMultiSel";
