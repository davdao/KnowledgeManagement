import * as React from 'react';
import { useState, useEffect } from 'react';
import { ComboBox, IComboBox, IComboBoxProps, Label } from 'office-ui-fabric-react';
import styles from './MultiComboBoxSelector.module.scss';

const LOADING_KEY = "LOADING_ITEM";

const MultiComboBoxSelector = (props) => {
    
    //const [itemDataFormated, setItemDataFormated] = useState(null);

    const [selectedOptionKeys, setSelectedOptionKeys] = useState([""]);
    const [textDisplayValue, setTextDisplayValue] = useState("");
    const [optionsList, setOptionsList] = useState(null);
    const [errorMessage, setErrorMessage] = useState("");
    const [searchInProgress, setSearchInProgress] = useState(false);    

    let comboRef = React.createRef<IComboBox>();

    useEffect(() => {
        //selectedoption
        getOption();
        props.onLoadOptions("").then((result) => {
            setOptionsList(result);
        })
        setTextDisplayValue(getTextDisplayValue());
    }, []);

    let comboProps: IComboBoxProps = {
        componentRef: comboRef,
        text: textDisplayValue,                 
        label: props.label,
        allowFreeform: props.allowFreeform ? props.allowFreeform : false,
        autoComplete:'on',
        disabled: props.disabled,
        styles: {
            input: {
                backgroundColor: 'inherit'
            }
        },
        useComboBoxAsMenuWidth: true,                
        options: optionsList,
        placeholder: props.placeholder, 
       // onRenderOption: onRenderOption(),
    };

    if (!props.allowMultiSelect) {
        
    } else {
        comboProps.onChange = onChangeMulti;
        comboProps.selectedKey = selectedOptionKeys;
        comboProps.multiSelect = true;
   //     comboProps.onMenuOpen = getOptions;
    }

    return(<div className={styles.multiComboBox}>
            <ComboBox options={optionsList} />
            {
                errorMessage ?
                    <Label className={styles.errorMessage}>{errorMessage}</Label> : null
            }
            </div>);

    async function getOption() {
        let optionList = await props.onLoadOptions("textinput");

        console.log(optionList);
    }

    function getTextDisplayValue(): string {

        let initialValue: string = null;    
        if (props.textDisplayValue) {    
            initialValue = props.textDisplayValue;    
        } else {
            if (props.allowMultiSelect) {
                initialValue = props.defaultSelectedKeys ? props.defaultSelectedKeys.toString() : '';
            } else {
                if (props.defaultSelectedKey) {
                    initialValue = props.defaultSelectedKey;
                } else {
                    initialValue = textDisplayValue;
                }
            }
        }        
        return initialValue;
    }

    /**
     * Retrieves all available options from the provided method
     * @param inputText if 'searchAsYouType' flag is enabled, pass the curren combo box text
     */
    async function getOptions(inputText): Promise<void> {
      /*  // Case when user opens the dropdown multiple times on the same field
        if(options.length > 0 && !props.searchAsYouType) {
            return;
        }
        else {
            let options: IComboBoxOption[] = [];

            setOptions([
                            {
                                key: LOADING_KEY,
                                text: '',
                                disabled: true,
                                itemType: SelectableOptionMenuItemType.Header                        
                            } as IComboBoxOption
                        ]);
            
            if(props.onLoadOptions) {
                options = await props.onLoadOptions(inputText);
            }
        }*/
    }

    function onChange(){

    }

    function onChangeMulti(){
    
    }

    function onRenderOption(){

    }
};

export default MultiComboBoxSelector;