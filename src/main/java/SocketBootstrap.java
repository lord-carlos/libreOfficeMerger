import com.sun.star.bridge.UnoUrlResolver;
import com.sun.star.bridge.XUnoUrlResolver;
import com.sun.star.comp.helper.Bootstrap;
import com.sun.star.comp.helper.BootstrapException;
import com.sun.star.connection.ConnectionSetupException;
import com.sun.star.connection.NoConnectException;
import com.sun.star.lang.IllegalArgumentException;
import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.lib.util.NativeLibraryLoader;
import com.sun.star.text.XTextDocument;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;

import java.io.File;
import java.io.IOException;
import java.net.URL;
import java.net.URLClassLoader;
import java.util.Arrays;

/**
 * A bootstrap connector which establishes a connection to an Open Office socket server.
 */
public class SocketBootstrap {
    private static SocketBootstrap bootstrap;
    private Server server;
    final private TemplateConnection templateConnection;

    private SocketBootstrap() {
        if(Boolean.getBoolean("templatewriter.ooo.usesocketrange")) {
            templateConnection = new LocalSocketTemplateConnection();
        } else {
            templateConnection = new FreeSocketTemplateConnection();
        }
    }

    /**
     * Returns the SocketBootstrap.
     *
     * @return bootstrappen
     */
    public static SocketBootstrap getDefault() {
        if(bootstrap == null) {
            bootstrap = new SocketBootstrap();
        }
        return bootstrap;
    }

    /**
     * Bootstraps a connection the server in the provided folder.
     *
     * @param folder is the Open Office folder, e.g.. ../openoffice/program/
     * @return the context of the created bootstrap connection
     * @throws BootstrapException if bootstrap fails
     */
    public XComponentContext bootstrap(String folder) throws BootstrapException {
        kill();
        server = new Server(folder);
        return connect();
    }

    public void kill() {
        if(server != null) {
            server.kill();
        }
    }

    /**
     * Creates the physical connection to the server.
     *
     * @return the context of the connection.
     * @throws BootstrapException if bootstrap fails
     */
    public XComponentContext connect() throws BootstrapException {
        System.out.println("starting to connecting to LibreOffice server, connect()");
        XComponentContext xContext;
        try {
            XComponentContext xLocalContext = getLocalContext();
            XMultiComponentFactory xLocalServiceManager = xLocalContext.getServiceManager();
            if(xLocalServiceManager == null) {
                throw new BootstrapException();
            }

            XUnoUrlResolver xUrlResolver = UnoUrlResolver.create(xLocalContext);

            server.start();

            try {
                xContext = waitForOpenOffice(xUrlResolver);
            } catch (BootstrapException | RuntimeException | InterruptedException e) {
                System.out.println("soffice has not responded to our connection request. We assume that it is in a bad state and kills it. " + e);
                try {
                    Runtime.getRuntime()
                           .exec(new String[]{System.getenv("WINDIR") + "\\system32\\" + "taskkill.exe",
                                              "/T",
                                              "/F",
                                              "/IM",
                                              "soffice.bin",
                                              "/FI",
                                              "USERNAME eq " + System.getenv("USERNAME")});
                    Runtime.getRuntime().exec(new String[]{"cmd", "/c", "tskill", "soffice"});
                    //                    Runtime.getRuntime().exec("tskill soffice");
                    Thread.sleep(500);
                } catch (IOException | InterruptedException e1) {
                    System.out.println("Error closing soffice: " + e1.getMessage());
                }
                server.start();
                xContext = waitForOpenOffice(xUrlResolver);
            }
        } catch (RuntimeException | BootstrapException e) {
            throw e;
        } catch (Exception e) {
            throw new BootstrapException(e.getMessage(), e);
        }
        return xContext;
    }

    private XComponentContext waitForOpenOffice(XUnoUrlResolver xUrlResolver) throws BootstrapException, ConnectionSetupException, IllegalArgumentException, InterruptedException {
        XComponentContext xContext = null;
        // Vent på at OpenOffice starter
        System.out.println("waitForOpenOffice(), waiting for libre office to startup");
        int i = 0;
        do {
            try {
                xContext = getRemoteContext(xUrlResolver);
            } catch (NoConnectException ex) {
                if(i++ == 60) {
                    throw new BootstrapException(ex.toString());
                }
                Thread.sleep(500);
            } finally {
                System.out.println("waitForOpenOffice(), Waited for LibreOffice connection: 500 millis * " + i + " = " + (500 * i));
            }
        } while(xContext == null);
        return xContext;
    }

    /**
     * Disconnects and kills the connection to the server.
     *
     * @param textDocument this document will be closed
     */
    public void disconnect(XTextDocument textDocument) {
        try {
            closeDocument(textDocument);
        } catch (Exception e) {
            // ignore
        }
        templateConnection.close();
        System.out.println("templateConnection.close()");
    }

    private XComponentContext getLocalContext() throws Exception {
        XComponentContext xLocalContext = Bootstrap.createInitialComponentContext(null);
        if(xLocalContext == null) {
            throw new BootstrapException();
        }
        return xLocalContext;
    }

    private XComponentContext getRemoteContext(XUnoUrlResolver xUrlResolver) throws BootstrapException, ConnectionSetupException, IllegalArgumentException, NoConnectException {

        XComponentContext xContext;
        try {
            xContext = UnoRuntime.queryInterface(XComponentContext.class, xUrlResolver.resolve(templateConnection.getConnectString()));
        } catch (Exception e) {
            throw new BootstrapException(e);
        }
        if(xContext == null) {
            throw new BootstrapException();
        }
        return xContext;
    }

    private class Server {
        private Process process;
        private String folder;
        private String[] options = new String[]{"--nologo", "--nodefault", "--norestore", "--nocrashreport", "--nolockcheck"};

        /**
         * Instantierer en OpenOfficeServer
         *
         * @param folder OpenOffice program folderen
         */
        public Server(String folder) {
            this.folder = folder;
        }

        /**
         * Starter OpenOffice serveren
         *
         * @throws BootstrapException if some general Bootstrap problem occurs
         * @throws IOException        if we cannot starting the process due to IO problems
         */
        public void start() throws BootstrapException, IOException {
            System.out.println("starting the libeOffice server, start()");
            URLClassLoader loader = new URLClassLoader(new URL[]{new File(folder).toURI().toURL()});

            File fOffice = NativeLibraryLoader.getResource(loader, "soffice.exe");
            if(fOffice == null) {
                throw new BootstrapException("Bookplan Writer executable not foundt.");
            }

            String[] command = new String[options.length + 2];
            command[0] = fOffice.getPath();
            int i = 1;
            for(String opt : options) {
                command[i++] = opt;
            }
            command[command.length - 1] = templateConnection.getAcceptString();

            System.out.println("OpenOffice started with command: '" + Arrays.asList(command) + "'");
            process = Runtime.getRuntime().exec(command);
            System.out.println("exec(command) finished.");
        }

        /**
         * Stopper OpenOffice serveren
         */
        public void kill() {
            if(process != null) {
                System.out.println("kill() - destroying process.");
                process.destroy();
                process = null;
            }
        }
    }

    public interface TemplateConnection {
        String getConnectString();

        String getAcceptString();

        /**
         * Frees reserved resources. Must be called after the OpenOffice process is ended.
         */
        void close();
    }

    private abstract static class AbstractSocketTemplateConnection implements TemplateConnection {
        abstract protected int getPort();

        final public String getConnectString() {
            return "uno:socket,host=localhost,port=" + getPort() + ";urp;StarOffice.ComponentContext";
        }

        final public String getAcceptString() {
            return "-accept=socket,host=localhost,port=" + getPort() + ";urp;";
        }
    }

    private static class LocalSocketTemplateConnection extends AbstractSocketTemplateConnection {
        //        private LocalSocketSelector localSocketSelector;

        private LocalSocketTemplateConnection() {
            //            this.localSocketSelector = new LocalSocketSelector();
        }

        @Override
        protected int getPort() {
            //            return localSocketSelector.getPort();
            return 4398;
        }

        public void close() {
            //            this.localSocketSelector.close();
        }
    }

    private static class FreeSocketTemplateConnection extends AbstractSocketTemplateConnection {
        //        private FreeSocketSelector freeSocketSelector;

        public FreeSocketTemplateConnection() {
            //            this.freeSocketSelector = new FreeSocketSelector();
        }

        @Override
        protected int getPort() {
            return 4399;
            //            return freeSocketSelector.getPort();
        }

        public void close() {
            // not necessary for FreeSocketSelector
        }
    }

    /**
     * Method copied from dev guide: http://wiki.services.openoffice.org/wiki/Documentation/DevGuide/OfficeDev/Closing_Documents
     *
     * @param xDocument the document to close
     */
    public void closeDocument(XTextDocument xDocument) {
        // Check supported functionality of the document (model or controller).
        com.sun.star.frame.XModel xModel = UnoRuntime.queryInterface(com.sun.star.frame.XModel.class, xDocument);
        System.out.println("Trying to close Document. closeDocument()");
        if(xModel != null) {
            // It is a full featured office document.
            // Try to use close mechanism instead of a hard dispose().
            // But maybe such service is not available on this model.
            com.sun.star.util.XCloseable xCloseable = UnoRuntime.queryInterface(com.sun.star.util.XCloseable.class, xModel);

            if(xCloseable != null) {
                try {
                    // use close(boolean DeliverOwnership)
                    // The boolean parameter DeliverOwnership tells objects vetoing
                    // the close process that they may
                    // assume ownership if they object the closure by throwing a
                    // CloseVetoException
                    // Here we give up ownership. To be on the safe side, catch
                    // possible veto exception anyway.
                    xCloseable.close(true);
                    System.out.println("document closed, xCloseable.close(true);");
                } catch (com.sun.star.util.CloseVetoException ignored) {
                }
            }
            // If close is not supported by this model - try to dispose it.
            // But if the model disagree with a reset request for the modify
            // state
            // we shouldn't do so. Otherwise some strange things can happen.
            else {
                com.sun.star.lang.XComponent xDisposeable = UnoRuntime.queryInterface(com.sun.star.lang.XComponent.class, xModel);
                xDisposeable.dispose();
                System.out.println("document disposed, xDisposeable.dispose();");
            }
        }
    }
}
