export type EmailNotificationEntity = {
  readonly type: 'emailNotification';
  readonly id: string;
  readonly conversationId?: string;
  readonly htmlBody?: string;
};
