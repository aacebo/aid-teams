import {
  AdaptiveCard,
  TextBlock,
  ChoiceSetInput,
  ActionSet,
  ExecuteAction,
  SubmitData,
} from '@microsoft/teams.cards';

export type PollOption = { title: string; value: string };

export function createPollCard(question: string, options: PollOption[]): AdaptiveCard {
  return new AdaptiveCard(
    new TextBlock(question, { wrap: true, weight: 'Bolder', size: 'Medium' }),
    new ChoiceSetInput(...options)
      .withId('choice')
      .withLabel('Select an option')
      .withIsRequired(true)
      .withErrorMessage('Please select an option before submitting.'),
    new ActionSet(
      new ExecuteAction({ title: 'Submit' })
        .withData(new SubmitData('poll.submit'))
        .withAssociatedInputs('auto')
        .withStyle('positive')
    )
  );
}
