import * as React from 'react';
import styles from './FuelForm.module.scss';
import { IFuelFormProps } from './IFuelFormProps';
import { IFuelFormState } from './IFuelFormState';
//import { escape } from '@microsoft/sp-lodash-subset';

import { spfi, SPFI, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";


import { Checkbox, ChoiceGroup, DefaultButton, IChoiceGroupOption } from 'office-ui-fabric-react';
import { Dropdown, IDropdownOption } from 'office-ui-fabric-react';
import { TextField, MaskedTextField } from 'office-ui-fabric-react';
import { DatePicker, DayOfWeek } from 'office-ui-fabric-react';
import { PeoplePicker, PrincipalType } from '@pnp/spfx-controls-react/lib/PeoplePicker';

import { IPersonaProps } from 'office-ui-fabric-react';

import { Stack, IStackTokens} from '@fluentui/react/lib/Stack';

const personChoiceGroupOption: IChoiceGroupOption[] = [
  { key: 'me', text: 'dla siebie' },
  { key: 'someone', text: 'dla innej osoby' },
]

const defaultStackToken: IStackTokens = {
  childrenGap: 10,
  padding: 10,
}

export default class FuelForm extends React.Component<IFuelFormProps, IFuelFormState> {
  private sp: SPFI;

  constructor(props: IFuelFormProps) {
    super(props); 
    
    this.state = {
      formForUser: true,
      driver: undefined,
      supervisor: undefined,
      supplier: undefined,
      mask: undefined,
      limitChange: 25,
      distance: 10,
      startDate: new Date(),
      endDate: new Date(),
      personOnList: true,
    }

    this.sp = spfi().using(SPFx(props.context));
    this.sp
  }

  

  public render(): React.ReactElement<IFuelFormProps> {

    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName,
      context,
      spolkiNames,
      dostawcy,
    } = this.props;

    const {
      formForUser,
      personOnList,
      driver,
      name,
      surname,
      email,
      company,
      registrationNumber,
      supplier,
      mask,
      cardNumber,
      distance,
      limitChange,
      route,
      startDate,
      endDate,
      justification,
      supervisor
    } = this.state;

    if (endDate < startDate) {
      this.setState({
        endDate: startDate
      })
    }
    
    if (supplier) {
      const supplierMask = dostawcy.filter(s => s.Title==supplier)[0].maska;
      if (mask!=supplierMask){
        this.setState({
          mask: supplierMask
        })
      }
    }

    //const expression: RegExp = /^[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}$/i;

    driver
    name
    surname
    email
    company
    registrationNumber
    supplier
    cardNumber
    distance
    limitChange
    route
    startDate
    endDate
    justification
    supervisor

    description
    isDarkTheme
    environmentMessage
    userDisplayName
    hasTeamsContext
    context

    const spolkiOptions: IDropdownOption[] = spolkiNames.map(n => ({key: n, text: n}));
    const dostawcyOptions: IDropdownOption[] = dostawcy.map(n => ({key: n.Title, text: n.Title}));

    let distanceMessage: string = '';
    let limitChangeMessage: string = '';

    if (distance<10) {
      distanceMessage = 'Dystans musi się zawierać w wartościach pomiędzy 10 a 9999'
    }
    if (limitChange%25!=0) {
      limitChangeMessage = 'Zmiana limitu musi być wielokrotnością 25'
    }

    return (
      <section className={`${styles.fuelForm} ${hasTeamsContext ? styles.teams : ''}`}>
        <Stack tokens={defaultStackToken}>

          <div>Wnioskujący: <strong>{context.pageContext.user.displayName}</strong></div>

          <Stack horizontal tokens={defaultStackToken} verticalAlign='center'>
            <Stack tokens={defaultStackToken}>
              <ChoiceGroup options={personChoiceGroupOption} defaultSelectedKey='me' onChange={this.onChangePersonChoice} label='Dla kogo składany jest wniosek'/>
              { !formForUser ?
                <Checkbox label='Osoby nie ma na liscie' onChange={this.onChangePersonNotOnList}/>
                : 
                null
              }
            </Stack>
            { !formForUser&&!personOnList ?
                <div>
                  <Stack horizontal>
                    <TextField label='Imię' onChange={this.onChangeName}/>
                    <TextField label='Nazwisko' onChange={this.onChangeSurname}/>
                  </Stack>
                  <TextField label='adres email' onChange={this.onChangeEmail}/>
                </div> 
              : null
            }
            { !formForUser&&personOnList ? 
              <PeoplePicker
                context={context as any}
                titleText="Wybierz osobę"
                personSelectionLimit={1}
                showtooltip={false}
                required={false}              
                onChange={this.onChangeDriver}
                showHiddenInUI={false}
                principalTypes={[PrincipalType.User]}
                resolveDelay={0} 
                ensureUser={true}
              />
              : null
            }
          </Stack>
          
          {/**/}
          <Stack horizontal tokens={defaultStackToken}>
            <PeoplePicker
                context={context as any}
                titleText="Wybierz przełożonego"
                personSelectionLimit={1}
                showtooltip={false}
                required={false}              
                onChange={this.onChangeSupervisor}
                showHiddenInUI={false}
                principalTypes={[PrincipalType.User]}
                resolveDelay={1000} 
                ensureUser={true}
            />
            <Dropdown
              label='Wybierz spółkę użytkującą spółkę'
              options={spolkiOptions}
            />

          </Stack>
          
          {/**/}
          <Stack horizontal tokens={defaultStackToken}>
            
            <TextField
              label='Wpisz numer rejestracyjny'
            />
            <Dropdown
              label='Wybierz dostawcę paliwa'
              options={dostawcyOptions}
              onChange={this.onChangeSupplier}
            />
            { mask||mask=='' ?
              <MaskedTextField
                label='Wpisz numer karty'
                mask={mask}
                onChange={this.onChangeCardNumber}
              />
              :
              <TextField
                label='Wpisz numer karty'
                onChange={this.onChangeCardNumber}
                disabled={mask === undefined}
              />
            }
          </Stack>

          {/**/}
          <Stack horizontal tokens={defaultStackToken}>
            <MaskedTextField
              label='Wpisz odległość podróży'
              mask='9999'
              value={distance.toString()}
              maskChar=''
              errorMessage={distanceMessage}
              onChange={this.onChangeDistance}
            />
            <MaskedTextField
              label='Wpisz wnioskowaną zmianę limitu'
              mask='999'
              value={limitChange.toString()}
              maskChar=''
              errorMessage={limitChangeMessage}
              onChange={this.onChangeLimit}
            />

            {/*<SpinButton
              label='Wpisz odległość podróży (w km)'
              labelPosition={Position.top}
              min={10}
              max={9999}
              step={1}
              incrementButtonAriaLabel="Zwiększ o 1"
              decrementButtonAriaLabel="Zmniejsz 1"
              
            />
            
            <SpinButton
              label='Wpisz wnioskowaną zmianę limitu'
              labelPosition={Position.top}
              min={25}
              max={999}
              step={25}
              incrementButtonAriaLabel="Zwiększ o 25"
              decrementButtonAriaLabel="Zmniejsz 25"
            />*/}
          </Stack>
          
          {/**/}
          <TextField
            label='Wpisz trasę przejazdu'
            multiline
            rows={2}
            resizable={false}
          />
          <Stack horizontal tokens={defaultStackToken}>
            <DatePicker
              label='Wprowadź datę rozpoczęcia'
              firstDayOfWeek={DayOfWeek.Monday}
              minDate={new Date()}
              value={startDate}
              onSelectDate={this.onChangeStartDate}
            />
            <DatePicker
              label='Wprowadź datę zakończenia'
              firstDayOfWeek={DayOfWeek.Monday}
              value={endDate}
              minDate={startDate}
              onSelectDate={this.onChangeEndDate}
            />
          </Stack>
          <TextField
              label='Wpisz uzasadnienie'
              multiline
              rows={3}
              resizable={false}
          />
          
        </Stack>
        <DefaultButton
          text='Złóż wniosek'
          onClick={() => {this.onClickSubmit(this.state)}}
        />
      </section>
    );
  }

  private onChangePersonNotOnList = (ev: React.FormEvent<HTMLElement | HTMLInputElement>, isChecked: boolean) => {
    this.setState({
      personOnList: !isChecked
    })
  }

  private onChangePersonChoice = (ev: React.FormEvent<HTMLElement | HTMLInputElement>, option: IChoiceGroupOption) => {
    this.setState({
      formForUser: option.key == 'me'
    })
  }

  private onChangeDriver = (items: IPersonaProps[]) => {
    this.setState({
      driver: items
    });
  }

  private onChangeName = (ev: React.FormEvent<HTMLElement | HTMLTextAreaElement>, name?: string) => {
    this.setState({
      name: name
    })
  }

  private onChangeSurname = (ev: React.FormEvent<HTMLElement | HTMLTextAreaElement>, surname?: string) => {
    this.setState({
      surname: surname
    })
  }

  private onChangeEmail = (ev: React.FormEvent<HTMLElement | HTMLTextAreaElement>, email?: string) => {
    this.setState({
      email: email
    })
  }

  private onChangeSupervisor = (items: IPersonaProps[]) => {
    this.setState({
      supervisor: items
    });
  }

  private onChangeSupplier = (ev: React.FormEvent<HTMLDivElement>, option: IDropdownOption) => {
    this.setState({
      supplier: option.text
    })
  }

  private onChangeDistance = (ev: React.FormEvent<HTMLElement | HTMLTextAreaElement>, distance?: string) => {
    let dst: number = parseInt(distance)
    this.setState({
      distance: dst
    })
  }

  private onChangeLimit = (ev: React.FormEvent<HTMLElement | HTMLTextAreaElement>, limit?: string) => {
    let limitChange: number = parseInt(limit)
    this.setState({
      limitChange: limitChange
    })
  }

  private onChangeCardNumber = (ev: React.FormEvent<HTMLElement | HTMLTextAreaElement>, cardNumber?: string) => {
    this.setState({
      cardNumber: cardNumber
    })
  }

  private onChangeStartDate = (date: Date) => {
    this.setState({
      startDate: date
    });
  }

  private onChangeEndDate = (date: Date) => {
    this.setState({
      endDate: date
    });
  }

  private onClickSubmit(state: IFuelFormState) {
    alert("Dziękuję za złozenie wniosku")
  }
}
