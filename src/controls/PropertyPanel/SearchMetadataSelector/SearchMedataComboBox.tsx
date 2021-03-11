import * as React from 'react';
import { useState, useEffect } from 'react';
import { IComboBox, IComboBoxOption } from 'office-ui-fabric-react';

/*************************************** */
//Sample Code from : PnP Modern Searchhttps://microsoft-search.github.io/pnp-modern-search/
/*************************************** */
const LOADING_KEY = "LOADING_ITEM";

const SearchMedataComboBox = (props) => {
    
    //const [itemDataFormated, setItemDataFormated] = useState(null);

    const [selectedOptionKeys, setSelectedOptionKeys] = useState([""]);
    const [textDisplayValue, setTextDisplayValue] = useState("");
    const [options, setOptions] = useState(null);
    const [errorMessage, setErrorMessage] = useState("");
    const [searchInProgress, setSearchInProgress] = useState(false);    

    let comboRef = React.createRef<IComboBox>();

    useEffect(() => {
        //selectedoption
        //avaiableoptions
        setTextDisplayValue(getTextDisplayValue());
    }, []);

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
};

function getOptions(){

}

function onChange(){

}

function onChangeMulti(){
   
}

function onRenderOption(){

}



export default SearchMedataComboBox;