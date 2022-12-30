import * as React from "react";
import styles from "./FuelForm.module.scss";
import { IFuelFormProps } from "./IFuelFormProps";
import { IFuelFormState } from "./IFuelFormState";

import { spfi, SPFI, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";

import {
  Checkbox,
  ChoiceGroup,
  DefaultButton,
  IChoiceGroupOption,
} from "office-ui-fabric-react";

import { Dropdown, IDropdownOption } from "office-ui-fabric-react";
import { TextField, MaskedTextField } from "office-ui-fabric-react";
import { DatePicker, DayOfWeek } from "office-ui-fabric-react";
import {
  PeoplePicker,
  PrincipalType,
} from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { Slider } from "office-ui-fabric-react";
import { MessageBar, MessageBarType } from "office-ui-fabric-react";

import { IPersonaProps } from "office-ui-fabric-react";

import { Stack, IStackTokens } from "@fluentui/react/lib/Stack";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { round } from "@microsoft/sp-lodash-subset";

const personChoiceGroupOption: IChoiceGroupOption[] = [
  { key: "me", text: "dla siebie" },
  { key: "someone", text: "dla innej osoby" },
];

const defaultStackToken: IStackTokens = {
  childrenGap: 10,
};

export interface IFormRecord {
  Title: string;
  Imie: string; // name
  Nazwisko: string; //surname
  email: string;
  spolka: string; // company
  Nr_rej: string; // registrationNumber
  DostawcaId: number; // supplier
  WBId: number;
  karta: string; // cardNumber
  dystans: number; // distance
  limit_x002b_: number; // limitChange
  Trasa: string; // route
  Data_od: Date; // starDate
  Data_do: Date; // endDate
  uzasadnienie: string; // justification
}

export interface ISupplier {
  "@odata.type": string;
  Id: number;
  Value: string;
}

const sliderValueFormat = (value: number): string => `${value} l`;

interface IBarProps {
  onDismiss?: () => void;
  message: string;
}

const SuccessBar = (p: IBarProps): JSX.Element => {
  return (
    <MessageBar messageBarType={MessageBarType.success} onDismiss={p.onDismiss}>
      {p.message}
    </MessageBar>
  );
};

const ErrorBar = (p: IBarProps): JSX.Element => {
  return (
    <MessageBar messageBarType={MessageBarType.error} onDismiss={p.onDismiss}>
      {p.message}
    </MessageBar>
  );
};

export default class FuelForm extends React.Component<
  IFuelFormProps,
  IFuelFormState
> {
  private sp: SPFI;

  constructor(props: IFuelFormProps) {
    super(props);

    this.state = {
      formForUser: true,
      driver: undefined,
      supervisor: undefined,
      supplier: undefined,
      mask: undefined,
      limitChange: 0,
      distance: 0.1,
      personOnList: true,
    };

    this.sp = spfi().using(SPFx(props.context));
  }

  public render(): React.ReactElement<IFuelFormProps> {
    const { hasTeamsContext, context, companyNames, suppliers } = this.props;

    const {
      formForUser,
      personOnList,
      success,
      formSent,
      showErrorBar,
      showErrorMessages,
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
      supervisor,
    } = this.state;

    if (endDate < startDate) {
      this.setState({
        endDate: startDate,
      });
    }

    if (supplier) {
      const supplierMask = suppliers.filter((s) => s.Id === supplier)[0].maska;
      if (mask !== supplierMask) {
        this.setState({
          mask: supplierMask,
        });
      }
    }

    const expression: RegExp = /^[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}$/i;

    let spolkiOptions: IDropdownOption[];
    if (companyNames) {
      spolkiOptions = companyNames.map((n) => ({ key: n, text: n }));
    }
    let dostawcyOptions: IDropdownOption[];
    if (suppliers) {
      dostawcyOptions = suppliers.map((n) => ({ key: n.Id, text: n.Title }));
    }

    let distanceErrorMessage: string = "";
    let emailErrorMessage: string = "";
    const defaultErrorMessage = "Pole niewypełnione";

    if (distance < 10 && distance !== 0.1) {
      distanceErrorMessage =
        "Dystans musi się zawierać w wartościach pomiędzy 10 a 9999";
    }

    if (email && !expression.test(email)) {
      emailErrorMessage = "Niepoprawny email";
    }

    let sentButtonVisible: boolean;

    if (
      company &&
      registrationNumber &&
      supplier &&
      cardNumber &&
      distance >= 10 &&
      limitChange !== 0 &&
      route &&
      startDate &&
      endDate &&
      justification &&
      supervisor
    ) {
      sentButtonVisible = true;
    }
    if (sentButtonVisible && !formForUser) {
      if (personOnList) {
        sentButtonVisible = driver ? true : false;
      } else {
        sentButtonVisible = name && surname && email ? true : false;
      }
    }

    return (
      <section
        className={`${styles.fuelForm} ${hasTeamsContext ? styles.teams : ""}`}
      >
        <h1>{this.props.title}</h1>
        <Stack tokens={defaultStackToken}>
          <div>
            Wnioskujący: <strong>{context.pageContext.user.displayName}</strong>
          </div>

          <Stack horizontal tokens={defaultStackToken} verticalAlign="center">
            <Stack tokens={defaultStackToken}>
              <ChoiceGroup
                options={personChoiceGroupOption}
                defaultSelectedKey="me"
                onChange={this.onChangePersonChoice}
                label="Dla kogo składany jest wniosek"
              />
              {!formForUser ? (
                <Checkbox
                  label="Osoby nie ma na liscie"
                  onChange={this.onChangePersonNotOnList}
                />
              ) : null}
            </Stack>
            {!formForUser && !personOnList ? (
              <div>
                <Stack horizontal>
                  <TextField
                    label="Imię"
                    placeholder="Wprowadź Imię"
                    onChange={this.onChangeName}
                    errorMessage={
                      !name && showErrorMessages ? defaultErrorMessage : ""
                    }
                  />
                  <TextField
                    label="Nazwisko"
                    placeholder="Wprowadź Nazwisko"
                    onChange={this.onChangeSurname}
                    errorMessage={
                      !surname && showErrorMessages ? defaultErrorMessage : ""
                    }
                  />
                </Stack>
                <TextField
                  label="adres email"
                  placeholder="Wprowadź email"
                  onChange={this.onChangeEmail}
                  errorMessage={
                    !email && showErrorMessages
                      ? defaultErrorMessage
                      : emailErrorMessage
                  }
                />
              </div>
            ) : null}
            {!formForUser && personOnList ? (
              <PeoplePicker
                context={context as any}
                titleText="Osoba"
                personSelectionLimit={1}
                showtooltip={false}
                required={false}
                onChange={this.onChangeDriver}
                showHiddenInUI={false}
                principalTypes={[PrincipalType.User]}
                resolveDelay={0}
                ensureUser={true}
                placeholder="Wybierz osobę"
                errorMessage={
                  !driver && showErrorMessages ? defaultErrorMessage : ""
                }
              />
            ) : null}
          </Stack>

          {/**/}
          <Stack horizontal tokens={defaultStackToken}>
            <PeoplePicker
              context={context as any}
              titleText="Przełożony"
              personSelectionLimit={1}
              showtooltip={false}
              required={false}
              onChange={this.onChangeSupervisor}
              showHiddenInUI={false}
              principalTypes={[PrincipalType.User]}
              resolveDelay={1000}
              ensureUser={true}
              placeholder="Wybierz przełożonego"
              errorMessage={
                !supervisor && showErrorMessages ? defaultErrorMessage : ""
              }
            />
            <Dropdown
              label="Spółka                                                   "
              options={spolkiOptions}
              placeholder="Wybierz spółkę użytkująca samochód"
              onChange={this.onChangeCompany}
              errorMessage={
                !company && showErrorMessages ? defaultErrorMessage : ""
              }
            />
          </Stack>

          {/**/}
          <Stack horizontal tokens={defaultStackToken}>
            <Dropdown
              label="Dostawca paliwa"
              options={dostawcyOptions}
              onChange={this.onChangeSupplier}
              placeholder="Wybierz dostawcę paliwa"
              errorMessage={
                !supplier && showErrorMessages ? defaultErrorMessage : ""
              }
            />
            {mask || mask === "" ? (
              <MaskedTextField
                label="Numer karty"
                mask={mask}
                onChange={this.onChangeCardNumber}
                placeholder="Wpisz numer karty"
                errorMessage={
                  !cardNumber && showErrorMessages ? defaultErrorMessage : ""
                }
              />
            ) : (
              <TextField
                label="Numer karty"
                onChange={this.onChangeCardNumber}
                disabled={mask === undefined}
                placeholder="Wpisz numer karty"
                errorMessage={
                  !cardNumber && showErrorMessages ? defaultErrorMessage : ""
                }
              />
            )}
          </Stack>
          <Slider
            label="Dodatkowy limit"
            min={25}
            max={this.props.maxFuelLimit ? this.props.maxFuelLimit : 500}
            step={25}
            valueFormat={sliderValueFormat}
            snapToStep
            onChange={this.onChangeLimit}
          />
          <Stack horizontal tokens={defaultStackToken}>
            <MaskedTextField
              label="Odległość podróży"
              mask="9999"
              value={round(distance).toString()}
              maskChar=""
              errorMessage={
                distance === 0.1 && showErrorMessages
                  ? defaultErrorMessage
                  : distanceErrorMessage
              }
              onChange={this.onChangeDistance}
              placeholder="Wpisz odległość podróży"
            />
            <TextField
              label="Numer rejestracyjny"
              placeholder="Wpisz numer rejestracyjny"
              onChange={this.onChangeRegistrationNumber}
              errorMessage={
                !registrationNumber && showErrorMessages
                  ? defaultErrorMessage
                  : ""
              }
            />
          </Stack>
          <TextField
            label="Trasa przejazdu"
            placeholder="Wpisz trasę przejazdu"
            multiline
            rows={2}
            resizable={false}
            onChange={this.onChangeRoute}
            errorMessage={
              !route && showErrorMessages ? defaultErrorMessage : ""
            }
          />
          <Stack horizontal tokens={defaultStackToken}>
            <DatePicker
              label="Wprowadź datę rozpoczęcia"
              firstDayOfWeek={DayOfWeek.Monday}
              minDate={new Date()}
              value={startDate}
              onSelectDate={this.onChangeStartDate}
              isRequired={showErrorMessages && !startDate}
            />
            <DatePicker
              label="Wprowadź datę zakończenia"
              firstDayOfWeek={DayOfWeek.Monday}
              value={endDate}
              minDate={startDate}
              onSelectDate={this.onChangeEndDate}
              isRequired={showErrorMessages && !endDate}
            />
          </Stack>
          <TextField
            label="Uzasadnienie"
            placeholder="Wpisz uzasadnienie (numer delegacji)"
            multiline
            rows={3}
            resizable={false}
            onChange={this.onChangeJustifaction}
            errorMessage={
              !justification && showErrorMessages ? defaultErrorMessage : ""
            }
          />
          {formSent && success ? (
          <SuccessBar
            onDismiss={() => {
              this.setState({ formSent: false });
            }}
            message={"Wniosek złożony poprawnie. Dziękujemy!"}
          />
          ) : null}
          {formSent && !success ? (
            <ErrorBar
              onDismiss={() => {
                this.setState({ formSent: false });
              }}
              message={
                "Wykryto błąd przy składaniu wniosku! Spróbuj jeszcze raz. W przypadku dalszych problemów skontaktuj się z administracją."
              }
            />
          ) : null}
          {showErrorBar ? (
            <ErrorBar
              onDismiss={() => {
                this.setState({ showErrorBar: false });
              }}
              message={
                !startDate || !endDate
                  ? "Wykryto błąd przy składaniu wniosku! Wybierz datę rozpoczęcia i zakończenia i spróbuj ponownie."
                  : "Wykryto błąd przy składaniu wniosku! Upewnij się, że wszystkie pola są dobrze wypełnione."
              }
            />
          ) : null}
          
        </Stack>
        <br/>
        <DefaultButton
            text="Złóż wniosek"
            onClick={() => {
              this.onClickSubmit(this.state, context, sentButtonVisible);
            }}
          />
      </section>
    );
  }

  private onChangePersonNotOnList = (
    ev: React.FormEvent<HTMLElement | HTMLInputElement>,
    isChecked: boolean
  ): void => {
    this.setState({
      personOnList: !isChecked,
    });
  };

  private onChangePersonChoice = (
    ev: React.FormEvent<HTMLElement | HTMLInputElement>,
    option: IChoiceGroupOption
  ): void => {
    this.setState({
      formForUser: option.key === "me",
    });
  };

  private onChangeDriver = (items: IPersonaProps[]): void => {
    this.setState({
      driver: items,
    });
  };

  private onChangeName = (
    ev: React.FormEvent<HTMLElement | HTMLTextAreaElement>,
    name?: string
  ): void => {
    this.setState({
      name: name,
    });
  };

  private onChangeSurname = (
    ev: React.FormEvent<HTMLElement | HTMLTextAreaElement>,
    surname?: string
  ): void => {
    this.setState({
      surname: surname,
    });
  };

  private onChangeEmail = (
    ev: React.FormEvent<HTMLElement | HTMLTextAreaElement>,
    email?: string
  ): void => {
    this.setState({
      email: email,
    });
  };

  private onChangeSupervisor = (items: IPersonaProps[]): void => {
    this.setState({
      supervisor: items,
    });
  };

  private onChangeCompany = (
    ev: React.FormEvent<HTMLDivElement>,
    option: IDropdownOption
  ): void => {
    this.setState({
      company: option.text,
    });
  };

  private onChangeRegistrationNumber = (
    ev: React.FormEvent<HTMLElement | HTMLTextAreaElement>,
    number?: string
  ): void => {
    this.setState({
      registrationNumber: number,
    });
  };

  private onChangeSupplier = (
    ev: React.FormEvent<HTMLDivElement>,
    option: IDropdownOption
  ): void => {
    this.setState({
      supplier: option.key as number,
    });
  };

  private onChangeCardNumber = (
    ev: React.FormEvent<HTMLElement | HTMLTextAreaElement>,
    cardNumber?: string
  ): void => {
    this.setState({
      cardNumber: cardNumber,
    });
  };

  private onChangeDistance = (
    ev: React.FormEvent<HTMLElement | HTMLTextAreaElement>,
    distance?: string
  ): void => {
    this.setState({
      distance: parseInt(distance),
    });
  };

  private onChangeLimit = (value: number): void => {
    this.setState({
      limitChange: value,
    });
  };

  private onChangeRoute = (
    ev: React.FormEvent<HTMLElement | HTMLTextAreaElement>,
    route?: string
  ): void => {
    this.setState({
      route: route,
    });
  };

  private onChangeStartDate = (date: Date): void => {
    this.setState({
      startDate: date,
    });
  };

  private onChangeEndDate = (date: Date): void => {
    this.setState({
      endDate: date,
    });
  };

  private onChangeJustifaction = (
    ev: React.FormEvent<HTMLElement | HTMLTextAreaElement>,
    justification?: string
  ): void => {
    this.setState({
      justification: justification,
    });
  };

  private onClickSubmit(
    state: IFuelFormState,
    context: WebPartContext,
    formCorect: boolean
  ): void {
    this.setState({
      showErrorBar: true,
      showErrorMessages: true,
    });

    if (!formCorect) {
      return;
    }

    let name: string = "";
    let surname: string = "";
    let email: string = "";
    if (state.formForUser) {
      email = context.pageContext.user.email;
    }
    if (!state.formForUser && state.personOnList) {
      email = state.driver[0].secondaryText;
    }
    if (!state.formForUser && !state.personOnList) {
      name = state.name;
      surname = state.surname;
      email = state.email;
    }

    const record: IFormRecord = {
      Title: "Nowy Wniosek",
      Imie: name,
      Nazwisko: surname,
      email: email,
      spolka: state.company,
      Nr_rej: state.registrationNumber,
      karta: state.cardNumber,
      dystans: state.distance,
      DostawcaId: state.supplier,
      WBId: parseInt(state.supervisor[0].id),
      limit_x002b_: state.limitChange,
      Trasa: state.route,
      Data_od: state.startDate,
      Data_do: state.endDate,
      uzasadnienie: state.justification,
    };

    let sentCorrectly = true;
    this.sp.web.lists
      .getById(this.props.targetListId)
      .items.add(record)
      .catch((error) => {
        const errorMessage: string = error + "\n" + JSON.stringify(record);
        this.sp.web.lists
          .getByTitle("FuelFormErrors")
          .items.add({
            Title: context.pageContext.user.email,
            errorMessage: errorMessage,
          })
          .catch((error) => alert(error));
        sentCorrectly = false;
      });
    this.setState({
      formSent: true,
      success: sentCorrectly,
      showErrorBar: false,
    });
  }
}
