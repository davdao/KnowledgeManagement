import * as React from 'react';
import { useState, useEffect } from 'react';
import { IComboBoxOption } from 'office-ui-fabric-react/lib/components/ComboBox/ComboBox.types';
import styles from './MultiComboBoxSelector.module.scss';
import { ComboBox } from 'office-ui-fabric-react/lib/components/ComboBox/ComboBox';

const LOADING_KEY = 'LOADING_ITEM';

interface IMultiComboBoxSelector {
    availableOptions: Promise<IComboBoxOption[]>;
    disabled?: boolean;
    label?: string;
}

const MultiComboBoxSelector = (props: IMultiComboBoxSelector) => {
    const [allAvailableOptions, setAllAvailableOptions] = useState([]);
    const [selectedOptionKeys, setSelectedOptionKeys] = useState([]);
    const [options, setOptions] = useState([""]);
    const [textDisplayValue, setTextDisplayValue] = useState("");

    useEffect(() => {
        props.availableOptions.then((result) => {
            setAllAvailableOptions(result);
        })
    }, []);

    props.availableOptions.then((result) => {
        console.log(result);
    })
    
    return(
            <div className={styles.MultiComboBoxSelector}>
                <ComboBox options={allAvailableOptions} 
                multiSelect={true}/>
            </div>
        );
};

export default MultiComboBoxSelector;